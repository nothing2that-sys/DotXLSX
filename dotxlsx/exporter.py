"""Export converted pixel art frames to a macro-free .xlsx workbook."""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from dotxlsx.converter import PALETTE_CHARS, ConvertedResult
from dotxlsx.vba_builder import build_vba_project_bin


@dataclass(frozen=True)
class ExportStyle:
    """Visual styling for the Excel Canvas sheet."""

    bg_hex: str = "#0F172A"
    fg_hex: str = "#38BDF8"
    font_name: str = "Segoe UI Symbol"
    font_size: int = 9
    col_width: float = 1.7
    row_height: float = 16.0
    color_cell_width: float = 2.2
    color_cell_height: float = 14.0


def _norm_hex(hex_color: str) -> str:
    return "FF" + hex_color.lstrip("#").upper().zfill(6)[-6:]


def _rgb_to_hex(rgb: tuple[int, int, int] | list[int] | object) -> str:
    r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    return f"FF{r:02X}{g:02X}{b:02X}"


def _write_xlsm_with_vba(wb: openpyxl.Workbook, out_path: Path) -> None:
    """Save workbook to BytesIO and inject pure-Python vbaProject.bin as .xlsm."""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    vba_bin = build_vba_project_bin()

    with zipfile.ZipFile(buf, "r") as zin, zipfile.ZipFile(
        out_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                txt = data.decode("utf-8").replace(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                )
                if "vbaProject.bin" not in txt:
                    txt = txt.replace(
                        "</Types>",
                        '<Override PartName="/xl/vbaProject.bin" '
                        'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
                    )
                data = txt.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                txt = data.decode("utf-8")
                if "vbaProject.bin" not in txt:
                    txt = txt.replace(
                        "</Relationships>",
                        '<Relationship Id="rIdDotXLSXVba" '
                        'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                        'Target="vbaProject.bin"/></Relationships>',
                    )
                data = txt.encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("xl/vbaProject.bin", vba_bin)


def _setup_engine_and_data(
    wb: openpyxl.Workbook,
    data_lines: list[list[str]],
    cols: int,
    rows: int,
    progress_cb: Callable[[int, int], None] | None,
    use_vba_autoplay: bool = False,
    frame_delay_ms: int = 100,
) -> int:
    """Create Engine and Data sheets and configure calculation or VBA autoplay."""
    total_frames = len(data_lines)
    if not use_vba_autoplay:
        wb.calculation.iterate = True
        wb.calculation.iterateCount = 1
        wb.calculation.calcMode = "auto"

    ws_engine = wb.create_sheet("Engine")
    ws_data = wb.create_sheet("Data")

    ws_engine["A1"] = "Parameter"
    ws_engine["B1"] = "Value"
    ws_engine["C1"] = "Description"

    ws_engine["A2"] = "Tick"
    if use_vba_autoplay:
        ws_engine["B2"] = 0
        ws_engine["C2"] = "0-based frame index; updated automatically by VBA timer"
    else:
        ws_engine["B2"] = "=IF(B3=1,0,MOD(IFERROR(N(B2),0)+1+0*RAND(),B4))"
        ws_engine["C2"] = "0-based frame index; press F9 to advance (uses RAND() volatile trigger)"

    ws_engine["A3"] = "Reset"
    ws_engine["B3"] = 0
    ws_engine["C3"] = "Set to 1 and press F9 to reset to frame 0"

    ws_engine["A4"] = "TotalFrames"
    ws_engine["B4"] = total_frames
    ws_engine["C4"] = "Total animation frames"

    ws_engine["A5"] = "FrameDelayMs"
    ws_engine["B5"] = max(30, int(frame_delay_ms))
    ws_engine["C5"] = "Frame interval in milliseconds (used by .xlsm VBA autoplay)"

    ws_engine["A6"] = "GridCols"
    ws_engine["B6"] = cols
    ws_engine["A7"] = "GridRows"
    ws_engine["B7"] = rows

    ws_engine.column_dimensions["A"].width = 16
    ws_engine.column_dimensions["B"].width = 28
    ws_engine.column_dimensions["C"].width = 42

    data_row = 1
    for f_idx, frame_rows in enumerate(data_lines):
        for line in frame_rows:
            c_data = ws_data.cell(row=data_row, column=1, value=line)
            c_data.data_type = "s"
            data_row += 1
        if progress_cb:
            progress_cb(f_idx + 1, total_frames * 2)

    return total_frames * rows


def _add_canvas_hud(
    ws_canvas: openpyxl.worksheet.worksheet.Worksheet,
    cols: int,
    rows: int,
    use_vba: bool = False,
) -> None:
    hud_row = rows + 2
    hud_font = Font(name="Malgun Gothic", size=9, color="94A3B8")
    label = (
        "Auto-playing GIF (.xlsm) | Frame:"
        if use_vba
        else "Press F9 to play next frame | Frame:"
    )
    ws_canvas.cell(row=hud_row, column=1, value=label).font = hud_font
    hud_frame = ws_canvas.cell(
        row=hud_row,
        column=min(cols, 18),
        value='=Engine!B2+1&" / "&Engine!B4',
    )
    hud_frame.font = hud_font


def export_to_xlsx(
    result: ConvertedResult,
    output_path: str | Path,
    style: ExportStyle | None = None,
    progress_cb: Callable[[int, int], None] | None = None,
) -> Path:
    """Write ConvertedResult into a .xlsx (or .xlsm if output_path ends with .xlsm)."""
    if style is None:
        style = ExportStyle()

    out = Path(output_path)
    cols = result.cols
    rows = result.rows
    is_xlsm = out.suffix.lower() == ".xlsm"

    wb = openpyxl.Workbook()
    if is_xlsm:
        wb.code_name = "ThisWorkbook"
    ws_canvas = wb.active
    ws_canvas.title = "Canvas"
    ws_canvas.views.sheetView[0].showGridLines = False
    avg_delay = int(sum(result.durations_ms) / max(1, len(result.durations_ms)))

    if result.render_mode == "color_cell":
        _export_color_cell(
            wb, ws_canvas, result, style, progress_cb, use_vba=is_xlsm, frame_delay_ms=avg_delay
        )
    elif result.render_mode == "color_braille":
        _export_color_braille(ws_canvas, result, style, progress_cb)
    else:
        _export_monochrome_braille(
            wb, ws_canvas, result, style, progress_cb, use_vba=is_xlsm, frame_delay_ms=avg_delay
        )

    if is_xlsm:
        _write_xlsm_with_vba(wb, out)
    else:
        wb.save(out)
    return out


def _export_monochrome_braille(
    wb: openpyxl.Workbook,
    ws_canvas: openpyxl.worksheet.worksheet.Worksheet,
    result: ConvertedResult,
    style: ExportStyle,
    progress_cb: Callable[[int, int], None] | None,
    use_vba: bool = False,
    frame_delay_ms: int = 100,
) -> None:
    total_frames = len(result.braille_frames)
    if total_frames == 0:
        raise ValueError("No frames to export.")

    cols, rows = result.cols, result.rows
    for c in range(1, cols + 1):
        ws_canvas.column_dimensions[get_column_letter(c)].width = style.col_width
    for r in range(1, rows + 1):
        ws_canvas.row_dimensions[r].height = style.row_height

    bg_fill = PatternFill(
        fill_type="solid",
        start_color=_norm_hex(style.bg_hex),
        end_color=_norm_hex(style.bg_hex),
    )
    fg_font = Font(
        name=style.font_name,
        size=style.font_size,
        color=_norm_hex(style.fg_hex),
    )
    center_align = Alignment(horizontal="center", vertical="center")

    if total_frames == 1:
        frame_rows = result.braille_frames[0]
        for r in range(1, rows + 1):
            line = frame_rows[r - 1]
            for c in range(1, cols + 1):
                cell = ws_canvas.cell(row=r, column=c, value=line[c - 1])
                cell.fill = bg_fill
                cell.font = fg_font
                cell.alignment = center_align
            if progress_cb:
                progress_cb(r, rows)
    else:
        total_data_rows = _setup_engine_and_data(
            wb,
            result.braille_frames,
            cols,
            rows,
            progress_cb,
            use_vba_autoplay=use_vba,
            frame_delay_ms=frame_delay_ms,
        )
        data_range = f"Data!$A$1:$A${total_data_rows}"
        formula = f"=MID(INDEX({data_range},IFERROR(Engine!$B$2,0)*{rows}+ROW()),COLUMN(),1)"

        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                cell = ws_canvas.cell(row=r, column=c, value=formula)
                cell.fill = bg_fill
                cell.font = fg_font
                cell.alignment = center_align
            if progress_cb:
                progress_cb(total_frames + int((r / rows) * total_frames), total_frames * 2)

        _add_canvas_hud(ws_canvas, cols, rows, use_vba=use_vba)


def _export_color_braille(
    ws_canvas: openpyxl.worksheet.worksheet.Worksheet,
    result: ConvertedResult,
    style: ExportStyle,
    progress_cb: Callable[[int, int], None] | None,
) -> None:
    if not result.braille_frames or not result.cell_fg_colors:
        raise ValueError("No color Braille data to export.")

    cols, rows = result.cols, result.rows
    for c in range(1, cols + 1):
        ws_canvas.column_dimensions[get_column_letter(c)].width = style.col_width
    for r in range(1, rows + 1):
        ws_canvas.row_dimensions[r].height = style.row_height

    bg_fill = PatternFill(
        fill_type="solid",
        start_color=_norm_hex(style.bg_hex),
        end_color=_norm_hex(style.bg_hex),
    )
    center_align = Alignment(horizontal="center", vertical="center")
    font_cache: dict[str, Font] = {}

    frame_rows = result.braille_frames[0]
    fg_arr = result.cell_fg_colors[0]
    for r in range(1, rows + 1):
        line = frame_rows[r - 1]
        for c in range(1, cols + 1):
            hex_fg = _rgb_to_hex(fg_arr[r - 1, c - 1])
            font = font_cache.get(hex_fg)
            if font is None:
                font = Font(name=style.font_name, size=style.font_size, color=hex_fg)
                font_cache[hex_fg] = font
            cell = ws_canvas.cell(row=r, column=c, value=line[c - 1])
            cell.fill = bg_fill
            cell.font = font
            cell.alignment = center_align
        if progress_cb:
            progress_cb(r, rows)


def _export_color_cell(
    wb: openpyxl.Workbook,
    ws_canvas: openpyxl.worksheet.worksheet.Worksheet,
    result: ConvertedResult,
    style: ExportStyle,
    progress_cb: Callable[[int, int], None] | None,
    use_vba: bool = False,
    frame_delay_ms: int = 100,
) -> None:
    total_frames = len(result.rgb_frames)
    if total_frames == 0:
        raise ValueError("No color cell frames to export.")

    cols, rows = result.cols, result.rows
    for c in range(1, cols + 1):
        ws_canvas.column_dimensions[get_column_letter(c)].width = style.color_cell_width
    for r in range(1, rows + 1):
        ws_canvas.row_dimensions[r].height = style.color_cell_height

    if total_frames == 1:
        rgb_arr = result.rgb_frames[0]
        fill_cache: dict[str, PatternFill] = {}
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                hex_c = _rgb_to_hex(rgb_arr[r - 1, c - 1])
                fill = fill_cache.get(hex_c)
                if fill is None:
                    fill = PatternFill(fill_type="solid", start_color=hex_c, end_color=hex_c)
                    fill_cache[hex_c] = fill
                ws_canvas.cell(row=r, column=c).fill = fill
            if progress_cb:
                progress_cb(r, rows)
    else:
        total_data_rows = _setup_engine_and_data(
            wb,
            result.indexed_frames,
            cols,
            rows,
            progress_cb,
            use_vba_autoplay=use_vba,
            frame_delay_ms=frame_delay_ms,
        )
        data_range = f"Data!$A$1:$A${total_data_rows}"
        formula = f"=MID(INDEX({data_range},IFERROR(Engine!$B$2,0)*{rows}+ROW()),COLUMN(),1)"

        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                cell = ws_canvas.cell(row=r, column=c, value=formula)
                cell.number_format = ";;;"
            if progress_cb:
                progress_cb(total_frames + int((r / rows) * total_frames), total_frames * 2)

        last_col = get_column_letter(cols)
        cell_range = f"A1:{last_col}{rows}"
        for idx, rgb in enumerate(result.palette_rgb):
            hex_c = _rgb_to_hex(rgb)
            fill = PatternFill(fill_type="solid", start_color=hex_c, end_color=hex_c)
            rule = CellIsRule(
                operator="equal",
                formula=[f'"{PALETTE_CHARS[idx]}"'],
                stopIfTrue=True,
                fill=fill,
            )
            ws_canvas.conditional_formatting.add(cell_range, rule)

        _add_canvas_hud(ws_canvas, cols, rows, use_vba=use_vba)
