"""Image and GIF to Unicode Braille and Color Pixel Art converter."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

import numpy as np
from PIL import Image, ImageEnhance, ImageOps, ImageSequence

DitherMode = Literal["floyd_steinberg", "threshold"]
RenderMode = Literal["monochrome_braille", "color_braille", "color_cell"]

# Unicode Braille Patterns block starts at U+2800.
BRAILLE_BASE = 0x2800
PALETTE_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUV"  # 32-color index encoding

# Bit weights for a 2x4 subpixel block (row, col):
# (0,0)=1, (1,0)=2, (2,0)=4, (3,0)=64
# (0,1)=8, (1,1)=16, (2,1)=32, (3,1)=128
_BRAILLE_WEIGHTS = np.array(
    [
        [0x01, 0x08],
        [0x02, 0x10],
        [0x04, 0x20],
        [0x40, 0x80],
    ],
    dtype=np.uint8,
)


@dataclass(frozen=True)
class ConversionConfig:
    """Configuration for converting an image/GIF into Excel pixel art."""

    cols: int = 80
    rows: int = 45
    render_mode: RenderMode = "monochrome_braille"
    dither: DitherMode = "floyd_steinberg"
    threshold: int = 128
    invert: bool = False
    auto_contrast: bool = True
    contrast: float = 1.2
    sharpness: float = 1.5
    frame_step: int = 1
    max_frames: int = 60
    palette_colors: int = 32
    bg_hex: str = "#0F172A"
    fg_hex: str = "#38BDF8"

    @property
    def pixel_width(self) -> int:
        return self.cols if self.render_mode == "color_cell" else self.cols * 2

    @property
    def pixel_height(self) -> int:
        return self.rows if self.render_mode == "color_cell" else self.rows * 4


@dataclass
class LoadedMedia:
    """Loaded still image or animated GIF frames as RGBA PIL Images."""

    path: Path
    frames: list[Image.Image]
    durations_ms: list[int]
    original_size: tuple[int, int]

    @property
    def is_animated(self) -> bool:
        return len(self.frames) > 1


@dataclass
class ConvertedResult:
    """Converted pixel art frames ready for preview and Excel export."""

    render_mode: RenderMode
    cols: int
    rows: int
    durations_ms: list[int]
    braille_frames: list[list[str]] = field(default_factory=list)
    bitmaps: list[np.ndarray] = field(default_factory=list)
    cell_fg_colors: list[np.ndarray] = field(default_factory=list)  # (rows, cols, 3) uint8
    rgb_frames: list[np.ndarray] = field(default_factory=list)      # (H, W, 3) uint8
    palette_rgb: list[tuple[int, int, int]] = field(default_factory=list)
    indexed_frames: list[list[str]] = field(default_factory=list)   # [frame][row] -> hex string


def load_media(path: str | Path) -> LoadedMedia:
    """Load a still image or animated GIF, compositing GIF frames properly."""
    p = Path(path)
    with Image.open(p) as img:
        original_size = img.size
        frames: list[Image.Image] = []
        durations: list[int] = []

        if getattr(img, "is_animated", False) and getattr(img, "n_frames", 1) > 1:
            canvas = Image.new("RGBA", img.size, (0, 0, 0, 0))
            for frame in ImageSequence.Iterator(img):
                rgba = frame.convert("RGBA")
                canvas = Image.alpha_composite(canvas.copy(), rgba)
                frames.append(canvas.copy())
                dur = int(frame.info.get("duration", img.info.get("duration", 100)) or 100)
                durations.append(max(20, dur))
        else:
            frames.append(img.convert("RGBA"))
            durations.append(100)

    return LoadedMedia(
        path=p,
        frames=frames,
        durations_ms=durations,
        original_size=original_size,
    )


def compute_aspect_rows(
    orig_width: int,
    orig_height: int,
    cols: int,
    render_mode: RenderMode = "monochrome_braille",
    cell_aspect: float = 0.5,
) -> int:
    """Compute grid rows for `cols` to preserve image aspect ratio."""
    if orig_width <= 0 or orig_height <= 0 or cols <= 0:
        return 45
    if render_mode == "color_cell":
        # Each cell is 1x1 square pixel in Excel.
        return max(1, round(cols * (orig_height / orig_width)))
    pixel_w = cols * 2
    target_pixel_h = pixel_w * (orig_height / orig_width) * (cell_aspect * 2.0)
    return max(1, round(target_pixel_h / 4.0))


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    s = hex_color.lstrip("#")
    if len(s) != 6:
        return (15, 23, 42)
    return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _composite_rgb(frame: Image.Image, bg_rgb: tuple[int, int, int]) -> Image.Image:
    """Composite RGBA frame onto a solid RGB background."""
    bg = Image.new("RGBA", frame.size, (*bg_rgb, 255))
    return Image.alpha_composite(bg, frame).convert("RGB")


def _enhance_rgb(img: Image.Image, config: ConversionConfig) -> Image.Image:
    """Apply auto-contrast, contrast, and sharpness before downsampling."""
    out = img
    if config.auto_contrast:
        out = ImageOps.autocontrast(out, cutoff=1)
    if abs(config.contrast - 1.0) > 0.05:
        out = ImageEnhance.Contrast(out).enhance(config.contrast)
    if abs(config.sharpness - 1.0) > 0.05:
        out = ImageEnhance.Sharpness(out).enhance(config.sharpness)
    return out


def _binarize(gray: Image.Image, dither: DitherMode, threshold: int, invert: bool) -> np.ndarray:
    """Convert grayscale image to boolean 2D array (True = dot ON)."""
    arr = np.asarray(gray, dtype=np.int16)

    if threshold != 128:
        shift = 128 - threshold
        arr = np.clip(arr + shift, 0, 255)

    if dither == "floyd_steinberg":
        shifted_img = Image.fromarray(arr.astype(np.uint8), mode="L")
        bw = shifted_img.convert("1", dither=Image.Dither.FLOYDSTEINBERG)
        bitmap = np.asarray(bw, dtype=bool)
    else:
        bitmap = arr >= 128

    if invert:
        bitmap = ~bitmap
    return bitmap


def bitmap_to_braille_rows(bitmap: np.ndarray) -> list[str]:
    """Encode a boolean (rows*4, cols*2) bitmap into Unicode Braille strings."""
    h, w = bitmap.shape
    if h % 4 != 0 or w % 2 != 0:
        raise ValueError(f"Bitmap shape {bitmap.shape} must be divisible by (4, 2)")

    rows = h // 4
    cols = w // 2
    blocks = bitmap.reshape(rows, 4, cols, 2).transpose(0, 2, 1, 3)
    masks = np.tensordot(blocks.astype(np.uint8), _BRAILLE_WEIGHTS, axes=([2, 3], [0, 1]))

    lines: list[str] = []
    for r in range(rows):
        row_chars = [chr(BRAILLE_BASE + int(val)) for val in masks[r]]
        lines.append("".join(row_chars))
    return lines


def _compute_block_fg_colors(
    rgb_arr: np.ndarray,
    bitmap: np.ndarray,
    rows: int,
    cols: int,
    fallback_rgb: tuple[int, int, int],
) -> np.ndarray:
    """Compute average RGB of ON pixels in each 2x4 cell block."""
    # Shape: (rows, cols, 4, 2, 3) and mask: (rows, cols, 4, 2)
    rgb_blocks = rgb_arr.reshape(rows, 4, cols, 2, 3).transpose(0, 2, 1, 3, 4)
    mask_blocks = bitmap.reshape(rows, 4, cols, 2).transpose(0, 2, 1, 3)

    weights = mask_blocks.astype(np.float32)[..., None]  # (rows, cols, 4, 2, 1)
    counts = weights.sum(axis=(2, 3))                    # (rows, cols, 1)
    sums = (rgb_blocks.astype(np.float32) * weights).sum(axis=(2, 3))

    out = np.empty((rows, cols, 3), dtype=np.uint8)
    valid = counts[..., 0] > 0
    out[valid] = np.clip(sums[valid] / counts[valid], 0, 255).astype(np.uint8)
    out[~valid] = fallback_rgb
    return out


def _render_color_braille_preview(
    bitmap: np.ndarray,
    cell_colors: np.ndarray,
    bg_rgb: tuple[int, int, int],
) -> np.ndarray:
    """Render (rows*4, cols*2, 3) preview matching cell-level Braille colors."""
    rows, cols, _ = cell_colors.shape
    expanded = np.repeat(np.repeat(cell_colors, 4, axis=0), 2, axis=1)
    out = np.empty_like(expanded)
    out[:, :] = bg_rgb
    out[bitmap] = expanded[bitmap]
    return out


def _convert_color_cell(
    sampled_frames: list[Image.Image],
    durations: list[int],
    config: ConversionConfig,
) -> ConvertedResult:
    """Convert frames to full-color (still) or 16-color quantized (GIF) cells."""
    bg_rgb = _hex_to_rgb(config.bg_hex)
    target_size = (config.cols, config.rows)
    resized_rgb = [
        _enhance_rgb(_composite_rgb(f, bg_rgb), config).resize(
            target_size, Image.Resampling.LANCZOS
        )
        for f in sampled_frames
    ]

    if len(resized_rgb) == 1:
        arr = np.asarray(resized_rgb[0], dtype=np.uint8)
        return ConvertedResult(
            render_mode="color_cell",
            cols=config.cols,
            rows=config.rows,
            durations_ms=durations,
            rgb_frames=[arr],
        )

    # Multi-frame GIF: build a shared palette (up to 32 colors) for Excel CF.
    n_colors = max(2, min(32, config.palette_colors))
    sample_step = max(1, len(resized_rgb) // 8)
    samples = resized_rgb[::sample_step][:8]
    sheet = Image.new("RGB", (config.cols * len(samples), config.rows))
    for i, s_img in enumerate(samples):
        sheet.paste(s_img, (i * config.cols, 0))

    palette_img = sheet.quantize(colors=n_colors, method=Image.Quantize.MEDIANCUT)
    raw_pal = palette_img.getpalette() or []
    palette_rgb: list[tuple[int, int, int]] = []
    for i in range(n_colors):
        base = i * 3
        if base + 2 < len(raw_pal):
            palette_rgb.append((raw_pal[base], raw_pal[base + 1], raw_pal[base + 2]))
        else:
            palette_rgb.append((0, 0, 0))

    dither_opt = (
        Image.Dither.FLOYDSTEINBERG
        if config.dither == "floyd_steinberg"
        else Image.Dither.NONE
    )

    rgb_frames: list[np.ndarray] = []
    indexed_frames: list[list[str]] = []
    for img in resized_rgb:
        q_img = img.quantize(palette=palette_img, dither=dither_opt)
        idx_arr = np.clip(np.asarray(q_img, dtype=np.uint8), 0, n_colors - 1)
        rgb_frames.append(np.asarray(q_img.convert("RGB"), dtype=np.uint8))

        frame_rows: list[str] = []
        for r in range(config.rows):
            frame_rows.append("".join(PALETTE_CHARS[int(v)] for v in idx_arr[r]))
        indexed_frames.append(frame_rows)

    return ConvertedResult(
        render_mode="color_cell",
        cols=config.cols,
        rows=config.rows,
        durations_ms=durations,
        rgb_frames=rgb_frames,
        palette_rgb=palette_rgb,
        indexed_frames=indexed_frames,
    )


def convert_media(media: LoadedMedia, config: ConversionConfig) -> ConvertedResult:
    """Convert loaded media frames according to the chosen render mode."""
    step = max(1, config.frame_step)
    sampled_indices = list(range(0, len(media.frames), step))
    if config.max_frames > 0:
        sampled_indices = sampled_indices[: config.max_frames]

    sampled_frames = [media.frames[idx] for idx in sampled_indices]
    durations = [media.durations_ms[idx] * step for idx in sampled_indices]

    if config.render_mode == "color_cell":
        return _convert_color_cell(sampled_frames, durations, config)

    # Braille modes (monochrome_braille or color_braille)
    target_size = (config.pixel_width, config.pixel_height)
    bg_rgb = _hex_to_rgb(config.bg_hex)
    fg_rgb = _hex_to_rgb(config.fg_hex)
    comp_bg = (255, 255, 255) if config.invert else (0, 0, 0)

    braille_frames: list[list[str]] = []
    bitmaps: list[np.ndarray] = []
    cell_fg_colors: list[np.ndarray] = []
    rgb_frames: list[np.ndarray] = []

    for frame in sampled_frames:
        rgb_img = _enhance_rgb(_composite_rgb(frame, comp_bg), config).resize(
            target_size, Image.Resampling.LANCZOS
        )
        gray = rgb_img.convert("L")
        bitmap = _binarize(gray, config.dither, config.threshold, config.invert)
        braille_frames.append(bitmap_to_braille_rows(bitmap))
        bitmaps.append(bitmap)

        if config.render_mode == "color_braille":
            rgb_arr = np.asarray(rgb_img, dtype=np.uint8)
            c_colors = _compute_block_fg_colors(
                rgb_arr, bitmap, config.rows, config.cols, fallback_rgb=fg_rgb
            )
            # Quantize to <= 64 colors so Excel never hits its unique-font limit.
            c_colors = np.asarray(
                Image.fromarray(c_colors, mode="RGB")
                .quantize(colors=64, method=Image.Quantize.MEDIANCUT)
                .convert("RGB"),
                dtype=np.uint8,
            )
            cell_fg_colors.append(c_colors)
            rgb_frames.append(_render_color_braille_preview(bitmap, c_colors, bg_rgb))

    return ConvertedResult(
        render_mode=config.render_mode,
        cols=config.cols,
        rows=config.rows,
        durations_ms=durations,
        braille_frames=braille_frames,
        bitmaps=bitmaps,
        cell_fg_colors=cell_fg_colors,
        rgb_frames=rgb_frames,
    )


def render_preview_image(
    result: ConvertedResult,
    frame_idx: int,
    bg_hex: str = "#0F172A",
    fg_hex: str = "#38BDF8",
) -> Image.Image:
    """Render a converted frame to an RGB PIL Image for preview."""
    if result.render_mode == "color_cell":
        return Image.fromarray(result.rgb_frames[frame_idx], mode="RGB")

    if result.render_mode == "color_braille" and result.rgb_frames:
        # Re-apply current bg_hex cleanly around ON dots
        bitmap = result.bitmaps[frame_idx]
        c_colors = result.cell_fg_colors[frame_idx]
        arr = _render_color_braille_preview(bitmap, c_colors, _hex_to_rgb(bg_hex))
        return Image.fromarray(arr, mode="RGB")

    bitmap = result.bitmaps[frame_idx]
    bg_rgb = _hex_to_rgb(bg_hex)
    fg_rgb = _hex_to_rgb(fg_hex)
    h, w = bitmap.shape
    rgb = np.empty((h, w, 3), dtype=np.uint8)
    rgb[:, :] = bg_rgb
    rgb[bitmap] = fg_rgb
    return Image.fromarray(rgb, mode="RGB")
