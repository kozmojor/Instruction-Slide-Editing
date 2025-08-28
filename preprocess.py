import argparse
import io
import os
from pathlib import Path

import requests
from datasets import load_dataset
from PIL import Image
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# -----------------------------
# PPTX building (pure python-pptx)
# -----------------------------
def build_pptx(items, out_path, layout="two_column", caption_max_chars=280):
    """
    Build a PPTX where each element in `items` is a dict:
      {"image": PIL.Image.Image, "caption": str}

    layout:
      - "two_column": left = image (~55%), right = caption
      - "title_content": title from first 8 words; below = image+caption area
    """
    prs = Presentation()
    # 16:9 canvas
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    SLIDE_W = prs.slide_width
    SLIDE_H = prs.slide_height
    MARGIN = Inches(0.4)

    for ex in items:
        img = ex["image"]  # PIL.Image.Image
        cap = (ex.get("caption") or "").strip()

        # Light truncation for readability
        if caption_max_chars and len(cap) > caption_max_chars:
            cap = cap[: caption_max_chars - 1] + "…"

        slide = prs.slides.add_slide(blank)

        if layout == "two_column":
            # Left = image (~55%), Right = caption
            left_col_w = int(SLIDE_W * 0.55)
            left_x = MARGIN
            left_y = MARGIN
            left_w = left_col_w
            left_h = SLIDE_H - MARGIN * 2

            right_x = left_x + left_w + Inches(0.1)
            right_y = MARGIN
            right_w = SLIDE_W - right_x - MARGIN
            right_h = SLIDE_H - MARGIN * 2

            # Add image
            bio = io.BytesIO()
            img.convert("RGB").save(bio, format="JPEG", quality=90)
            bio.seek(0)
            slide.shapes.add_picture(bio, left_x, left_y, width=left_w, height=left_h)

            # Add caption
            tb = slide.shapes.add_textbox(right_x, right_y, right_w, right_h)
            tf = tb.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = cap
            run.font.size = Pt(20)
            p.alignment = PP_ALIGN.LEFT

        elif layout == "title_content":
            # Title (top), then image+caption area
            title_x = MARGIN
            title_y = MARGIN
            title_w = SLIDE_W - 2 * MARGIN
            title_h = Inches(0.9)

            body_x = MARGIN
            body_y = title_y + title_h + Inches(0.2)
            body_w = SLIDE_W - 2 * MARGIN
            body_h = SLIDE_H - body_y - MARGIN

            # Title = first 8 words
            title_text = " ".join(cap.split()[:8]) or "COCO Slide"
            tb_title = slide.shapes.add_textbox(title_x, title_y, title_w, title_h)
            tf = tb_title.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = title_text
            r.font.size = Pt(32)

            # Body: image left (60%), caption right (40%)
            img_w = int(body_w * 0.6)
            bio = io.BytesIO()
            img.convert("RGB").save(bio, format="JPEG", quality=90)
            bio.seek(0)
            slide.shapes.add_picture(bio, body_x, body_y, width=img_w, height=body_h)

            cap_x = body_x + img_w + Inches(0.2)
            cap_w = body_w - img_w - Inches(0.2)
            tb = slide.shapes.add_textbox(cap_x, body_y, cap_w, body_h)
            tf2 = tb.text_frame
            tf2.clear()
            p2 = tf2.paragraphs[0]
            r2 = p2.add_run()
            r2.text = cap
            r2.font.size = Pt(18)

        else:
            raise ValueError(f"Unknown layout: {layout!r}")

    prs.save(out_path)
    return out_path


# -----------------------------
# Dataset loader (Hugging Face)
# -----------------------------
def load_and_project_dataset(split: str, count: int, timeout: int, skip_broken: bool):
    """
    Load a tiny slice from `yerevann/coco-karpathy` and return:
        [{"image": PIL.Image.Image, "caption": str}, ...]
    Robust to sentences/captions being list[str] or list[dict].
    """
    import io, requests
    from datasets import load_dataset
    from PIL import Image

    # map "validation" -> "val"
    hf_split = "train" if split == "train" else "val"
    ds = load_dataset("yerevann/coco-karpathy", split=f"{hf_split}[:{count}]")

    def extract_caption(ex):
        # 1) sentences: list of dict OR list of str
        if "sentences" in ex and ex["sentences"]:
            s0 = ex["sentences"][0]
            if isinstance(s0, dict):
                return s0.get("raw") or s0.get("caption") or s0.get("text") or ""
            elif isinstance(s0, str):
                return s0
        # 2) captions: list of dict OR list of str
        if "captions" in ex and ex["captions"]:
            c0 = ex["captions"][0]
            if isinstance(c0, dict):
                return c0.get("caption") or c0.get("raw") or c0.get("text") or ""
            elif isinstance(c0, str):
                return c0
        # 3) single fields
        return ex.get("caption") or ex.get("text") or ""

    def extract_url(ex):
        return ex.get("coco_url") or ex.get("image_url") or ex.get("url")

    items = []
    for ex in ds:
        cap = (extract_caption(ex) or "").strip()
        url = extract_url(ex)
        if not url:
            if skip_broken:
                continue
            raise ValueError("No image URL field in example.")

        try:
            resp = requests.get(url, timeout=timeout)
            resp.raise_for_status()
            img = Image.open(io.BytesIO(resp.content)).convert("RGB")
        except Exception as e:
            if skip_broken:
                continue
            raise RuntimeError(f"Failed to fetch image from {url}: {e}") from e

        items.append({"image": img, "caption": cap})

    return items


# -----------------------------
# CLI
# -----------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--count", type=int, default=100, help="number of examples to include")
    ap.add_argument("--split", type=str, default="train", choices=["train", "validation"])
    ap.add_argument("--seed", type=int, default=42, help="(kept for interface parity; not used with server-side slicing)")
    ap.add_argument("--layout", type=str, default="two_column", choices=["two_column", "title_content"])
    ap.add_argument("--caption_max_chars", type=int, default=280)
    ap.add_argument("--timeout", type=int, default=20, help="HTTP timeout (seconds) for image download")
    ap.add_argument("--skip_broken", action="store_true", help="skip samples when image download fails")
    ap.add_argument("--out", type=str, default="coco_raw.pptx")
    args = ap.parse_args()

    # 1) Load a tiny slice of a COCO-like captions dataset (with URLs) and map it to {image: PIL, caption: str}
    items = load_and_project_dataset(
        split=args.split,
        count=args.count,
        timeout=args.timeout,
        skip_broken=args.skip_broken,
    )

    # 2) Build a raw PPTX (each slide: one image + one caption)
    out_path = Path(args.out).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    path = build_pptx(
        items=items,
        out_path=str(out_path),
        layout=args.layout,
        caption_max_chars=args.caption_max_chars,
    )
    print(f"[OK] Saved: {path}")


if __name__ == "__main__":
    main()
