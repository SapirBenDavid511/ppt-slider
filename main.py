#!/usr/bin/env python3
import argparse
import json
import math
from pathlib import Path
from typing import List, Tuple

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.dml.color import RGBColor
from PIL import Image
import os
import urllib.parse


SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}


def to_file_uri(p: Path) -> str:
    """
    Return a file:// URI that PowerPoint will accept.
    Handles local absolute paths and UNC paths on Windows.
    """
    try:
        rp = p.resolve()
    except Exception:
        rp = p

    if os.name == "nt":
        s = str(rp)

        # UNC path? (\\server\share\path)
        if s.startswith("\\\\"):
            # file:////server/share/path (4 slashes after 'file:')
            # Convert backslashes to forward slashes and percent-encode specials
            unc = s.replace("\\", "/")
            return "file:" + urllib.parse.urlsplit("////" + unc.lstrip("/")).geturl()

        # Local drive (e.g., C:\path\to\file.pdf)
        # as_uri() works fine for absolute paths with drive letters
        if rp.is_absolute():
            return rp.as_uri()

        # Fallback: make absolute then as_uri
        return rp.absolute().as_uri()
    else:
        # POSIX/macOS/Linux
        return rp.absolute().as_uri()

def parse_args():
    p = argparse.ArgumentParser(description="Build a PowerPoint from images and a config.json")
    p.add_argument("--config", type=str, default="config.json", help="Path to config.json")
    p.add_argument("--images-dir", type=str, default="images", help="Directory containing image files")
    p.add_argument("--pdf-dir", type=str, default="pdf-files", help="Directory containing PDF files")
    p.add_argument("--output", type=str, default="slides.pptx", help="Output PPTX filename")
    p.add_argument("--width-in", type=float, default=13.333, help="Slide width in inches (16:9 ~ 13.333x7.5)")
    p.add_argument("--height-in", type=float, default=7.5, help="Slide height in inches")
    return p.parse_args()

def load_config(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    return cfg

def numeric_sort_keys(keys: List[str]) -> List[str]:
    def key_func(k):
        try:
            return (0, int(k))
        except:
            return (1, k)
    return sorted(keys, key=key_func)

def add_title_and_subtitle(slide, title_text: str, subtitle_text: str, slide_width,
                           top_margin_in=0.25, side_margin_in=0.5, max_title_h_in=0.8, max_subtitle_h_in=0.55):
    # Title
    left = Inches(side_margin_in)
    top = Inches(top_margin_in)
    width = Inches(slide_width - 2*side_margin_in)
    height = Inches(max_title_h_in)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_frame = title_box.text_frame
    title_frame.clear()
    p = title_frame.paragraphs[0]
    run = p.add_run()
    run.text = title_text or ""
    run.font.size = Pt(34)
    run.font.bold = True
    p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

    # Subtitle (smaller gap if empty)
    sub_top = top_margin_in + (0.0 if not title_text else max_title_h_in * 0.75)
    subtitle_box = slide.shapes.add_textbox(Inches(side_margin_in), Inches(sub_top), width, Inches(max_subtitle_h_in))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.clear()
    p2 = subtitle_frame.paragraphs[0]
    run2 = p2.add_run()
    run2.text = subtitle_text or ""
    run2.font.size = Pt(18)
    run2.font.color.rgb = RGBColor(90, 90, 90)
    p2.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

    used_top_in = sub_top + (0.0 if not subtitle_text else max_subtitle_h_in * 0.8)
    return used_top_in

def choose_grid(n:int, slide_w_in: float, slide_h_in: float, max_rows:int=4, max_cols:int=4) -> Tuple[int,int]:
    """
    Choose rows x cols for up to 12 images balancing:
    - minimal empty cells
    - grid aspect approximates slide aspect (wider than tall usually)
    Constraints: rows<=max_rows, cols<=max_cols, rows*cols >= n
    """
    assert n >= 1
    best = None
    slide_aspect = slide_w_in / slide_h_in
    for rows in range(1, min(max_rows, n)+1):
        cols = math.ceil(n / rows)
        if cols > max_cols:
            continue
        capacity = rows * cols
        empty = capacity - n
        grid_aspect = cols / rows
        aspect_penalty = abs(grid_aspect - slide_aspect)
        score = (empty * 10.0) + aspect_penalty  # empty cells are heavily penalized
        # Prefer more columns for wide slides if tie
        tie_break = -cols
        cand = (score, tie_break, rows, cols)
        if best is None or cand < best:
            best = cand
    if best is None:
        # fallback: clamp to max rows/cols
        rows = min(max_rows, n)
        cols = min(max_cols, math.ceil(n/rows))
        return rows, cols
    return best[2], best[3]

def fit_rect_keep_aspect(img_w_px: int, img_h_px: int, cell_w_in: float, cell_h_in: float, dpi: int = 96) -> Tuple[float, float]:
    w_in_native = img_w_px / dpi
    h_in_native = img_h_px / dpi
    scale = min(cell_w_in / w_in_native, cell_h_in / h_in_native)
    return (max(0.01, w_in_native * scale), max(0.01, h_in_native * scale))

def resolve_pdf_paths(pdfs, pdf_dir: Path):
    if not pdfs:
        return []
    if isinstance(pdfs, str):
        pdfs = [pdfs]
    out = []
    for name in pdfs:
        p = Path(name)
        if not p.is_absolute():
            p = pdf_dir / name
        if p.exists() and p.suffix.lower() == ".pdf":
            out.append(p)
        else:
            print(f"[WARN] Missing/unsupported pdf: {p}")
    return out

def place_pdf_icons(slide, pdf_paths, canvas_left_in, canvas_top_in, canvas_w_in, canvas_h_in):
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

    max_per_col = max(8, int(canvas_h_in // 0.5))
    col = 0
    row_in_col = 0
    icon_w_in = 0.5
    icon_h_in = 0.5
    gap_in = 0.18
    label_w_in = max(2.5, canvas_w_in/2.5)

    for idx, p in enumerate(pdf_paths):
        if row_in_col >= max_per_col:
            col += 1
            row_in_col = 0
        x = canvas_left_in + col * (icon_w_in + label_w_in + 1.0)
        y = canvas_top_in + row_in_col * (icon_h_in + gap_in)

        # Build BOTH a file:// URI and a raw absolute path (Windows UNC/drive friendly)
        try:
            abs_path = str(p.resolve())
        except Exception:
            abs_path = str(p)
        file_uri = to_file_uri(Path(abs_path))

        # --- ICON (rectangle) ---
        icon = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(icon_w_in), Inches(icon_h_in))
        tf = icon.text_frame
        tf.clear()
        para = tf.paragraphs[0]
        run_icon = para.add_run()
        run_icon.text = "PDF"
        run_icon.font.size = Pt(12)
        para.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

        # Apply hyperlink to the ICON using RAW PATH (works well on many Windows setups)
        icon.click_action.hyperlink.address = abs_path
        icon.click_action.hyperlink.screen_tip = abs_path

        # --- LABEL (filename) ---
        lbl = slide.shapes.add_textbox(Inches(x + icon_w_in + 0.15), Inches(y - 0.02),
                                       Inches(label_w_in), Inches(icon_h_in + 0.04))
        ltf = lbl.text_frame
        ltf.clear()
        lp = ltf.paragraphs[0]
        run_lbl = lp.add_run()
        run_lbl.text = p.name
        run_lbl.font.size = Pt(14)
        lp.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

        # Apply hyperlink to the LABEL RUN as a FILE URI (works well cross-platform)
        # (Link the run itself, not just the textbox shape)
        run_lbl.hyperlink.address = file_uri

        # Also put the same file URI on the label shape (belt & suspenders)
        lbl.click_action.hyperlink.address = file_uri
        lbl.click_action.hyperlink.screen_tip = abs_path  # shows readable path on hover

        # Debug so you can see exactly what was written
        print("[LINK-ICON-RAW ]", abs_path)
        print("[LINK-LABEL-URI]", file_uri)

        row_in_col += 1

def resolve_image_paths(images: List[str], images_dir: Path) -> List[Path]:
    out = []
    for name in images:
        p = Path(name)
        if not p.is_absolute():
            p = images_dir / name
        if p.exists() and p.suffix.lower() in SUPPORTED_EXTS:
            out.append(p)
        else:
            print(f"[WARN] Missing/unsupported image: {p}")
    return out

def place_images(slide, image_paths: List[Path],
                 canvas_left_in: float, canvas_top_in: float, canvas_w_in: float, canvas_h_in: float,
                 slide_w_in: float, slide_h_in: float, layout: str = None):
    n = len(image_paths)
    # Allow up to 12 images per slide
    if n > 12:
        print(f"[INFO] More than 12 images provided ({n}); only the first 12 will be placed on this slide.")
        image_paths = image_paths[:12]
        n = 12

    # Pick rows/cols smartly
    if n == 2 and layout in ("horizontal", "vertical"):
        rows, cols = (1, 2) if layout == "horizontal" else (2, 1)
    else:
        rows, cols = choose_grid(n, slide_w_in, slide_h_in, max_rows=4, max_cols=4)

    # Gutters scale according to density
    base_gutter = 0.25
    density = n / (rows*cols)
    gutter_in = max(0.12, base_gutter * (0.9 if density < 0.75 else 1.05))

    # Outer padding
    outer_pad_w = 0.0
    outer_pad_h = 0.0

    cell_w_in = (canvas_w_in - 2*outer_pad_w - (cols-1)*gutter_in) / cols
    cell_h_in = (canvas_h_in - 2*outer_pad_h - (rows-1)*gutter_in) / rows

    # Preload image sizes
    sizes = []
    for p in image_paths:
        try:
            with Image.open(p) as im:
                sizes.append((p, im.width, im.height))
        except Exception as e:
            print(f"[WARN] Could not open image: {p} ({e})")

    # Place images row by row; center last row if partially filled
    idx = 0
    for r in range(rows):
        remaining = n - idx
        if remaining <= 0:
            break
        images_this_row = min(cols, remaining)
        start_col = 0
        if images_this_row < cols:
            start_col = (cols - images_this_row) // 2

        for c in range(images_this_row):
            if idx >= len(sizes):
                break
            p, w_px, h_px = sizes[idx]
            idx += 1

            w_in, h_in = fit_rect_keep_aspect(w_px, h_px, cell_w_in, cell_h_in)

            col_index = start_col + c
            cell_left_in = canvas_left_in + outer_pad_w + col_index * (cell_w_in + gutter_in)
            cell_top_in  = canvas_top_in  + outer_pad_h + r * (cell_h_in + gutter_in)

            offset_left_in = cell_left_in + (cell_w_in - w_in) / 2.0
            offset_top_in  = cell_top_in  + (cell_h_in - h_in) / 2.0

            slide.shapes.add_picture(str(p), Inches(offset_left_in), Inches(offset_top_in),
                                     width=Inches(w_in), height=Inches(h_in))

def build_ppt(cfg: dict, images_dir: Path, pdf_dir: Path, output_path: Path, slide_w_in: float, slide_h_in: float):
    prs = Presentation()
    prs.slide_width = Inches(slide_w_in)
    prs.slide_height = Inches(slide_h_in)

    blank_layout = prs.slide_layouts[6]  # blank
    side_margin_in = 0.45
    bottom_margin_in = 0.35

    for key in numeric_sort_keys(list(cfg.keys())):
        item = cfg[key]
        title = item.get("title", "")
        subtitle = item.get("sub-title", "") or item.get("subtitle", "")
        images_list = item.get("images", []) or []
        pdf_field = item.get("pdf", [])

        slide = prs.slides.add_slide(blank_layout)
        used_top_in = add_title_and_subtitle(slide, title, subtitle, slide_w_in)

        # Compute canvas for images; reduce title block if there are many images
        count = max(len(images_list), len(pdf_field) if isinstance(pdf_field, list) else (1 if pdf_field else 0))
        extra_top_cut = 0.0
        if count >= 8:
            extra_top_cut = 0.2
        if count >= 10:
            extra_top_cut = 0.35

        canvas_left_in = side_margin_in
        canvas_top_in = max(used_top_in + 0.15 - extra_top_cut, 0.6)
        canvas_w_in = slide_w_in - 2*side_margin_in
        canvas_h_in = slide_h_in - canvas_top_in - bottom_margin_in

        pdf_paths = resolve_pdf_paths(pdf_field, pdf_dir)
        if pdf_paths:
            place_pdf_icons(slide, pdf_paths, canvas_left_in, canvas_top_in, canvas_w_in, canvas_h_in)
        else:
            image_paths = resolve_image_paths(images_list, images_dir)        
            if not image_paths:
                tb = slide.shapes.add_textbox(Inches(canvas_left_in), Inches(canvas_top_in),
                                            Inches(canvas_w_in), Inches(0.6))
                tf = tb.text_frame
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = "No images found for this slide."
                run.font.size = Pt(16)
                run.font.color.rgb = RGBColor(180, 0, 0)
            else:
                place_images(slide, image_paths, canvas_left_in, canvas_top_in, canvas_w_in, canvas_h_in,
                            slide_w_in, slide_h_in, layout=item.get("layout"))

    prs.save(str(output_path))
    print(f"[OK] Wrote {output_path}")

def main():
    args = parse_args()
    cfg = load_config(Path(args.config))
    build_ppt(cfg, Path(args.images_dir), Path(args.pdf_dir), Path(args.output), args.width_in, args.height_in)

if __name__ == "__main__":
    main()
