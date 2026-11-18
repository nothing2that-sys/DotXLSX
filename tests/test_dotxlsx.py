"""Unit and integration tests for DotXLSX converter and exporter."""

from __future__ import annotations

import zipfile
from pathlib import Path

import numpy as np
import openpyxl
from PIL import Image

from dotxlsx.converter import (
    BRAILLE_BASE,
    ConversionConfig,
    bitmap_to_braille_rows,
    compute_aspect_rows,
    convert_media,
    load_media,
    render_preview_image,
)
from dotxlsx.exporter import ExportStyle, export_to_xlsx


def test_braille_bitmask_all_eight_dots() -> None:
    expected_weights = [
        (0, 0, 0x01),
        (1, 0, 0x02),
        (2, 0, 0x04),
        (3, 0, 0x40),
        (0, 1, 0x08),
        (1, 1, 0x10),
        (2, 1, 0x20),
        (3, 1, 0x80),
    ]
    for r, c, weight in expected_weights:
        bmp = np.zeros((4, 2), dtype=bool)
        bmp[r, c] = True
        rows = bitmap_to_braille_rows(bmp)
        assert rows == [chr(BRAILLE_BASE + weight)]


def test_compute_aspect_rows_preserves_ratio() -> None:
    assert compute_aspect_rows(160, 180, cols=80, render_mode="monochrome_braille") == 45
    assert compute_aspect_rows(200, 100, cols=40, render_mode="monochrome_braille") == 10
    assert compute_aspect_rows(200, 100, cols=40, render_mode="color_cell") == 20


def test_still_image_monochrome_braille_export(tmp_path: Path) -> None:
    img_path = tmp_path / "still.png"
    Image.new("RGB", (64, 64), (255, 255, 255)).save(img_path)

    media = load_media(img_path)
    cfg = ConversionConfig(
        cols=20, rows=10, render_mode="monochrome_braille", dither="threshold", threshold=128
    )
    result = convert_media(media, cfg)
    assert result.braille_frames[0][0][0] == chr(0x28FF)

    out_xlsx = tmp_path / "still_mono.xlsx"
    export_to_xlsx(result, out_xlsx, ExportStyle(bg_hex="#0F172A", fg_hex="#38BDF8"))
    wb = openpyxl.load_workbook(out_xlsx)
    assert wb.sheetnames == ["Canvas"]
    assert wb["Canvas"]["A1"].value == chr(0x28FF)
    assert not wb["Canvas"].views.sheetView[0].showGridLines


def test_still_image_color_braille_export(tmp_path: Path) -> None:
    img_path = tmp_path / "red.png"
    Image.new("RGB", (40, 40), (255, 0, 0)).save(img_path)

    media = load_media(img_path)
    cfg = ConversionConfig(
        cols=10, rows=5, render_mode="color_braille", dither="threshold", threshold=50
    )
    result = convert_media(media, cfg)
    out_xlsx = tmp_path / "color_braille.xlsx"
    export_to_xlsx(result, out_xlsx)

    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb["Canvas"]
    assert ws["A1"].value == chr(0x28FF)
    assert ws["A1"].font.color.rgb.endswith("FF0000")


def test_still_image_color_cell_export(tmp_path: Path) -> None:
    img_path = tmp_path / "blue.png"
    Image.new("RGB", (20, 20), (0, 128, 255)).save(img_path)

    media = load_media(img_path)
    cfg = ConversionConfig(
        cols=10, rows=10, render_mode="color_cell", auto_contrast=False, contrast=1.0, sharpness=1.0
    )
    result = convert_media(media, cfg)

    out_xlsx = tmp_path / "color_still.xlsx"
    export_to_xlsx(result, out_xlsx)
    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb["Canvas"]
    assert ws["A1"].value is None
    assert ws["A1"].fill.start_color.rgb.endswith("0080FF")


def test_animated_gif_monochrome_braille_export(tmp_path: Path) -> None:
    gif_path = tmp_path / "anim.gif"
    f1 = Image.new("RGB", (40, 40), (0, 0, 0))
    f2 = Image.new("RGB", (40, 40), (255, 255, 255))
    f3 = Image.new("RGB", (40, 40), (0, 0, 0))
    f1.save(gif_path, save_all=True, append_images=[f2, f3], duration=80, loop=0)

    media = load_media(gif_path)
    cfg = ConversionConfig(
        cols=10, rows=5, render_mode="monochrome_braille", dither="threshold"
    )
    result = convert_media(media, cfg)

    out_xlsx = tmp_path / "anim_mono.xlsx"
    export_to_xlsx(result, out_xlsx)
    wb = openpyxl.load_workbook(out_xlsx)
    assert wb.sheetnames == ["Canvas", "Engine", "Data"]
    assert wb.calculation.iterate is True
    assert wb["Canvas"]["A1"].value == "=MID(INDEX(Data!$A$1:$A$15,IFERROR(Engine!$B$2,0)*5+ROW()),COLUMN(),1)"

    with zipfile.ZipFile(out_xlsx) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        assert 'iterate="1"' in wb_xml


def test_animated_gif_color_cell_export(tmp_path: Path) -> None:
    gif_path = tmp_path / "color_anim.gif"
    f1 = Image.new("RGB", (20, 20), (255, 0, 0))
    f2 = Image.new("RGB", (20, 20), (0, 255, 0))
    f1.save(gif_path, save_all=True, append_images=[f2], duration=100, loop=0)

    media = load_media(gif_path)
    cfg = ConversionConfig(
        cols=8, rows=8, render_mode="color_cell", dither="threshold", palette_colors=4
    )
    result = convert_media(media, cfg)
    assert len(result.indexed_frames) == 2
    assert len(result.palette_rgb) == 4

    out_xlsx = tmp_path / "anim_color.xlsx"
    export_to_xlsx(result, out_xlsx)
    wb = openpyxl.load_workbook(out_xlsx)
    assert wb.sheetnames == ["Canvas", "Engine", "Data"]
    assert wb.calculation.iterate is True
    ws = wb["Canvas"]
    assert ws["A1"].value == "=MID(INDEX(Data!$A$1:$A$16,IFERROR(Engine!$B$2,0)*8+ROW()),COLUMN(),1)"
    assert ws["A1"].number_format == ";;;"
    assert len(ws.conditional_formatting) > 0


def test_animated_gif_xlsm_autoplay_export(tmp_path: Path) -> None:
    gif_path = tmp_path / "auto.gif"
    f1 = Image.new("RGB", (20, 20), (255, 0, 0))
    f2 = Image.new("RGB", (20, 20), (0, 255, 0))
    f1.save(gif_path, save_all=True, append_images=[f2], duration=120, loop=0)

    media = load_media(gif_path)
    cfg = ConversionConfig(
        cols=8, rows=8, render_mode="color_cell", dither="threshold", palette_colors=4
    )
    result = convert_media(media, cfg)

    out_xlsm = tmp_path / "auto.xlsm"
    export_to_xlsx(result, out_xlsm)
    assert out_xlsm.exists()

    with zipfile.ZipFile(out_xlsm) as zf:
        names = set(zf.namelist())
        assert "xl/vbaProject.bin" in names
        ct = zf.read("[Content_Types].xml").decode("utf-8")
        assert "sheet.macroEnabled.main+xml" in ct
        assert "vbaProject.bin" in ct

    wb = openpyxl.load_workbook(out_xlsm, keep_vba=True)
    assert wb["Engine"]["B2"].value == 0
    assert wb["Engine"]["B4"].value == 2
    assert wb["Engine"]["B5"].value == 120


def test_render_preview_image_modes() -> None:
    img_path_bmp = np.array([[False, True], [True, False]], dtype=bool)
    from dotxlsx.converter import ConvertedResult

    res = ConvertedResult(
        render_mode="monochrome_braille",
        cols=1,
        rows=1,
        durations_ms=[100],
        bitmaps=[img_path_bmp],
    )
    img = render_preview_image(res, 0, bg_hex="#000000", fg_hex="#FFFFFF")
    assert img.size == (2, 2)
    assert img.getpixel((0, 0)) == (0, 0, 0)
    assert img.getpixel((1, 0)) == (255, 255, 255)
