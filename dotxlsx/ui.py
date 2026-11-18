"""Tkinter Windows desktop UI for DotXLSX."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

from PIL import Image, ImageTk

from dotxlsx.converter import (
    ConversionConfig,
    ConvertedResult,
    DitherMode,
    LoadedMedia,
    RenderMode,
    compute_aspect_rows,
    convert_media,
    load_media,
    render_preview_image,
)
from dotxlsx.exporter import ExportStyle, export_to_xlsx

RESOLUTION_PRESETS: dict[str, tuple[int, int] | None] = {
    "소형 (40×22)": (40, 22),
    "표준 (80×45)": (80, 45),
    "고해상도 (120×68)": (120, 68),
    "사용자 지정": None,
}

RENDER_MODE_LABELS: dict[str, RenderMode] = {
    "컬러 셀 (풀컬러 / GIF 32색)": "color_cell",
    "컬러 점자 (정적 고해상도)": "color_braille",
    "단색 점자 (고속 / GIF 최적)": "monochrome_braille",
}

THEME_PRESETS: dict[str, tuple[str, str]] = {
    "Cyber Blue": ("#0F172A", "#38BDF8"),
    "Game Boy": ("#0F380F", "#9BBC0F"),
    "Classic Dark": ("#000000", "#FFFFFF"),
    "Classic Light": ("#FFFFFF", "#111827"),
    "Matrix Green": ("#090D0A", "#22C55E"),
}


class DotXLSXApp:
    """Main desktop application window."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("DotXLSX - 이미지/GIF → 엑셀 픽셀아트 변환기")
        self.root.geometry("1160x780")
        self.root.minsize(940, 660)

        self.media: LoadedMedia | None = None
        self.converted: ConvertedResult | None = None
        self.current_frame_idx: int = 0
        self.is_playing: bool = False
        self._play_after_id: str | None = None
        self._convert_after_id: str | None = None
        self._preview_photo: ImageTk.PhotoImage | None = None
        self._updating_aspect: bool = False

        # Tk variables
        self.file_info_var = tk.StringVar(value="선택된 파일 없음")
        self.render_mode_label_var = tk.StringVar(value="컬러 셀 (풀컬러 / GIF 32색)")
        self.preset_var = tk.StringVar(value="표준 (80×45)")
        self.cols_var = tk.IntVar(value=80)
        self.rows_var = tk.IntVar(value=45)
        self.aspect_lock_var = tk.BooleanVar(value=True)
        self.res_summary_var = tk.StringVar(value="해상도: 80 × 45 px (3,600 셀)")

        self.dither_var = tk.StringVar(value="floyd_steinberg")
        self.threshold_var = tk.IntVar(value=128)
        self.threshold_label_var = tk.StringVar(value="128")
        self.invert_var = tk.BooleanVar(value=False)
        self.auto_contrast_var = tk.BooleanVar(value=True)
        self.contrast_var = tk.DoubleVar(value=1.2)
        self.contrast_label_var = tk.StringVar(value="1.2")
        self.sharpness_var = tk.DoubleVar(value=1.5)
        self.sharpness_label_var = tk.StringVar(value="1.5")

        self.frame_step_var = tk.IntVar(value=1)
        self.max_frames_var = tk.IntVar(value=60)
        self.palette_colors_var = tk.IntVar(value=32)

        self.theme_var = tk.StringVar(value="Cyber Blue")
        self.bg_hex_var = tk.StringVar(value="#0F172A")
        self.fg_hex_var = tk.StringVar(value="#38BDF8")

        self.show_original_var = tk.BooleanVar(value=False)
        self.frame_pos_var = tk.IntVar(value=0)
        self.frame_text_var = tk.StringVar(value="0 / 0")
        self.status_var = tk.StringVar(value="이미지 또는 GIF 파일을 불러오세요.")

        self._build_ui()
        self._bind_events()
        self._update_mode_controls()

    def _build_ui(self) -> None:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")

        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_frame = ttk.Frame(main_pane, width=350)
        main_pane.add(left_frame, weight=0)

        # 1. File Load
        f_box = ttk.LabelFrame(left_frame, text="1. 이미지 / GIF 불러오기", padding=8)
        f_box.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(f_box, text="파일 열기 (PNG / JPG / GIF)...", command=self._on_open_file).pack(
            fill=tk.X
        )
        ttk.Label(f_box, textvariable=self.file_info_var, foreground="#475569", wraplength=310).pack(
            anchor=tk.W, pady=(5, 0)
        )

        # 2. Render Mode & Resolution
        r_box = ttk.LabelFrame(left_frame, text="2. 렌더링 모드 및 해상도", padding=8)
        r_box.pack(fill=tk.X, pady=(0, 6))

        mode_cb = ttk.Combobox(
            r_box,
            textvariable=self.render_mode_label_var,
            values=list(RENDER_MODE_LABELS.keys()),
            state="readonly",
        )
        mode_cb.pack(fill=tk.X, pady=(0, 6))
        mode_cb.bind("<<ComboboxSelected>>", self._on_mode_changed)

        preset_cb = ttk.Combobox(
            r_box,
            textvariable=self.preset_var,
            values=list(RESOLUTION_PRESETS.keys()),
            state="readonly",
        )
        preset_cb.pack(fill=tk.X, pady=(0, 6))
        preset_cb.bind("<<ComboboxSelected>>", self._on_preset_selected)

        grid_row = ttk.Frame(r_box)
        grid_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(grid_row, text="가로(열):").pack(side=tk.LEFT)
        self.cols_spin = ttk.Spinbox(
            grid_row, from_=10, to=200, textvariable=self.cols_var, width=6
        )
        self.cols_spin.pack(side=tk.LEFT, padx=(4, 12))

        ttk.Label(grid_row, text="세로(행):").pack(side=tk.LEFT)
        self.rows_spin = ttk.Spinbox(
            grid_row, from_=10, to=150, textvariable=self.rows_var, width=6
        )
        self.rows_spin.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Checkbutton(
            r_box,
            text="원본 종횡비 자동 맞춤",
            variable=self.aspect_lock_var,
            command=self._on_aspect_toggle,
        ).pack(anchor=tk.W, pady=(2, 2))
        ttk.Label(r_box, textvariable=self.res_summary_var, foreground="#0284C7").pack(anchor=tk.W)

        # 3. Conversion / Dithering & Detail Enhancement
        c_box = ttk.LabelFrame(left_frame, text="3. 디테일 보정 및 디더링", padding=8)
        c_box.pack(fill=tk.X, pady=(0, 6))

        ac_row = ttk.Frame(c_box)
        ac_row.pack(fill=tk.X, pady=(0, 3))
        ttk.Checkbutton(
            ac_row,
            text="자동 대비 (Auto-Contrast)",
            variable=self.auto_contrast_var,
            command=self._schedule_convert,
        ).pack(side=tk.LEFT)
        self.invert_chk = ttk.Checkbutton(
            ac_row,
            text="반전",
            variable=self.invert_var,
            command=self._schedule_convert,
        )
        self.invert_chk.pack(side=tk.RIGHT)

        ct_row = ttk.Frame(c_box)
        ct_row.pack(fill=tk.X, pady=(1, 1))
        ttk.Label(ct_row, text="대비:", width=7).pack(side=tk.LEFT)
        ttk.Scale(
            ct_row,
            from_=0.5,
            to=2.5,
            orient=tk.HORIZONTAL,
            variable=self.contrast_var,
            command=self._on_contrast_slide,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Label(ct_row, textvariable=self.contrast_label_var, width=4).pack(side=tk.RIGHT)

        sh_row = ttk.Frame(c_box)
        sh_row.pack(fill=tk.X, pady=(1, 3))
        ttk.Label(sh_row, text="선명도:", width=7).pack(side=tk.LEFT)
        ttk.Scale(
            sh_row,
            from_=0.0,
            to=3.0,
            orient=tk.HORIZONTAL,
            variable=self.sharpness_var,
            command=self._on_sharpness_slide,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Label(sh_row, textvariable=self.sharpness_label_var, width=4).pack(side=tk.RIGHT)

        d_row = ttk.Frame(c_box)
        d_row.pack(fill=tk.X, pady=(0, 3))
        ttk.Radiobutton(
            d_row,
            text="Floyd-Steinberg",
            variable=self.dither_var,
            value="floyd_steinberg",
            command=self._schedule_convert,
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(
            d_row,
            text="Threshold / 없음",
            variable=self.dither_var,
            value="threshold",
            command=self._schedule_convert,
        ).pack(side=tk.LEFT)

        t_row = ttk.Frame(c_box)
        t_row.pack(fill=tk.X, pady=(1, 0))
        ttk.Label(t_row, text="임계값:", width=7).pack(side=tk.LEFT)
        self.threshold_scale = ttk.Scale(
            t_row,
            from_=1,
            to=254,
            orient=tk.HORIZONTAL,
            variable=self.threshold_var,
            command=self._on_threshold_slide,
        )
        self.threshold_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Label(t_row, textvariable=self.threshold_label_var, width=4).pack(side=tk.RIGHT)

        # 4. GIF Options
        self.gif_box = ttk.LabelFrame(left_frame, text="4. GIF 애니메이션 설정", padding=8)
        self.gif_box.pack(fill=tk.X, pady=(0, 6))

        g_row = ttk.Frame(self.gif_box)
        g_row.pack(fill=tk.X)
        ttk.Label(g_row, text="스킵:").pack(side=tk.LEFT)
        self.step_spin = ttk.Spinbox(
            g_row, from_=1, to=10, textvariable=self.frame_step_var, width=4
        )
        self.step_spin.pack(side=tk.LEFT, padx=(3, 8))

        ttk.Label(g_row, text="최대:").pack(side=tk.LEFT)
        self.max_frames_spin = ttk.Spinbox(
            g_row, from_=2, to=300, textvariable=self.max_frames_var, width=4
        )
        self.max_frames_spin.pack(side=tk.LEFT, padx=(3, 8))

        ttk.Label(g_row, text="팔레트:").pack(side=tk.LEFT)
        self.palette_spin = ttk.Spinbox(
            g_row, from_=2, to=32, textvariable=self.palette_colors_var, width=3
        )
        self.palette_spin.pack(side=tk.LEFT, padx=(3, 0))

        # 5. Theme / Background Colors
        self.th_box = ttk.LabelFrame(left_frame, text="5. 색상 설정", padding=8)
        self.th_box.pack(fill=tk.X, pady=(0, 6))

        self.theme_cb = ttk.Combobox(
            self.th_box,
            textvariable=self.theme_var,
            values=list(THEME_PRESETS.keys()),
            state="readonly",
        )
        self.theme_cb.pack(fill=tk.X, pady=(0, 6))
        self.theme_cb.bind("<<ComboboxSelected>>", self._on_theme_selected)

        clr_row = ttk.Frame(self.th_box)
        clr_row.pack(fill=tk.X)
        self.bg_swatch = tk.Label(clr_row, width=3, bg=self.bg_hex_var.get(), relief=tk.SOLID, bd=1)
        self.bg_swatch.pack(side=tk.LEFT, padx=(0, 4))
        self.bg_btn = ttk.Button(clr_row, text="배경색", command=self._pick_bg_color)
        self.bg_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.fg_swatch = tk.Label(clr_row, width=3, bg=self.fg_hex_var.get(), relief=tk.SOLID, bd=1)
        self.fg_swatch.pack(side=tk.LEFT, padx=(0, 4))
        self.fg_btn = ttk.Button(clr_row, text="도트색", command=self._pick_fg_color)
        self.fg_btn.pack(side=tk.LEFT)

        # 6. Export
        e_box = ttk.Frame(left_frame)
        e_box.pack(fill=tk.X, pady=(4, 0))
        self.export_btn = ttk.Button(
            e_box,
            text="엑셀 파일 내보내기 (.xlsx)",
            command=self._on_export,
            state=tk.DISABLED,
        )
        self.export_btn.pack(fill=tk.X, ipady=6)
        self.progress = ttk.Progressbar(e_box, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(e_box, textvariable=self.status_var, foreground="#334155", wraplength=320).pack(
            anchor=tk.W
        )

        # Right panel: Preview
        right_frame = ttk.Frame(main_pane)
        main_pane.add(right_frame, weight=1)

        top_bar = ttk.Frame(right_frame)
        top_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(top_bar, text="실시간 미리보기", font=("Malgun Gothic", 10, "bold")).pack(
            side=tk.LEFT
        )
        ttk.Checkbutton(
            top_bar,
            text="원본 이미지 보기",
            variable=self.show_original_var,
            command=self._render_current_frame,
        ).pack(side=tk.RIGHT)

        self.preview_canvas = tk.Canvas(
            right_frame, bg="#0F172A", highlightthickness=1, highlightbackground="#334155"
        )
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)

        play_bar = ttk.Frame(right_frame)
        play_bar.pack(fill=tk.X, pady=(6, 0))
        self.play_btn = ttk.Button(
            play_bar, text="▶ 재생", width=9, command=self._toggle_play, state=tk.DISABLED
        )
        self.play_btn.pack(side=tk.LEFT)

        self.frame_scale = ttk.Scale(
            play_bar,
            from_=0,
            to=0,
            orient=tk.HORIZONTAL,
            variable=self.frame_pos_var,
            command=self._on_frame_seek,
        )
        self.frame_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Label(play_bar, textvariable=self.frame_text_var, width=10).pack(side=tk.RIGHT)

        self._set_gif_controls_state(False)

    def _bind_events(self) -> None:
        self.cols_var.trace_add("write", self._on_cols_changed)
        self.rows_var.trace_add("write", self._on_rows_changed)
        self.frame_step_var.trace_add("write", lambda *_: self._schedule_convert())
        self.max_frames_var.trace_add("write", lambda *_: self._schedule_convert())
        self.palette_colors_var.trace_add("write", lambda *_: self._schedule_convert())
        self.preview_canvas.bind("<Configure>", lambda _e: self._render_current_frame())

    def _get_render_mode(self) -> RenderMode:
        return RENDER_MODE_LABELS.get(self.render_mode_label_var.get(), "color_cell")

    def _update_mode_controls(self) -> None:
        mode = self._get_render_mode()
        is_braille = mode != "color_cell"
        is_mono = mode == "monochrome_braille"

        self.threshold_scale.configure(state=tk.NORMAL if is_braille else tk.DISABLED)
        self.invert_chk.configure(state=tk.NORMAL if is_braille else tk.DISABLED)

        self.theme_cb.configure(state="readonly" if is_mono else tk.DISABLED)
        self.fg_btn.configure(state=tk.NORMAL if is_mono else tk.DISABLED)

        if is_mono:
            self.th_box.configure(text="5. 단색 테마 (배경색 + 도트색)")
        elif mode == "color_braille":
            self.th_box.configure(text="5. 배경색 (도트는 원본 컬러 사용)")
        else:
            self.th_box.configure(text="5. 배경색 (투명 이미지 여백용)")

        is_gif_color = mode == "color_cell" and self.media is not None and self.media.is_animated
        self.palette_spin.configure(state=tk.NORMAL if is_gif_color else tk.DISABLED)
        self._update_res_summary()

    def _on_mode_changed(self, _event=None) -> None:
        self._update_mode_controls()
        if self.aspect_lock_var.get() and self.media is not None:
            self._apply_aspect_rows()
        self._schedule_convert()

    def _set_gif_controls_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self.step_spin.configure(state=state)
        self.max_frames_spin.configure(state=state)
        self.play_btn.configure(state=state)
        self._update_mode_controls()

    def _on_open_file(self) -> None:
        path = filedialog.askopenfilename(
            title="이미지 또는 GIF 선택",
            filetypes=[
                ("이미지 및 GIF", "*.png *.jpg *.jpeg *.bmp *.webp *.gif"),
                ("GIF 애니메이션", "*.gif"),
                ("모든 파일", "*.*"),
            ],
        )
        if not path:
            return

        self._stop_playback()
        try:
            self.media = load_media(path)
        except Exception as exc:
            messagebox.showerror("파일 열기 실패", f"이미지를 불러오지 못했습니다:\n{exc}")
            return

        w, h = self.media.original_size
        n_frames = len(self.media.frames)
        kind = f"GIF ({n_frames} 프레임)" if self.media.is_animated else "정적 이미지"
        self.file_info_var.set(f"{Path(path).name}  |  {w}×{h} px  |  {kind}")

        self._set_gif_controls_state(self.media.is_animated)
        self.export_btn.configure(state=tk.NORMAL)

        if self.aspect_lock_var.get():
            self._apply_aspect_rows()

        self.current_frame_idx = 0
        self._run_convert()
        if self.media.is_animated:
            self._start_playback()

    def _on_preset_selected(self, _event=None) -> None:
        dims = RESOLUTION_PRESETS.get(self.preset_var.get())
        if dims is None:
            return
        cols, rows = dims
        self._updating_aspect = True
        self.cols_var.set(cols)
        if self.aspect_lock_var.get() and self.media is not None:
            w, h = self.media.original_size
            self.rows_var.set(compute_aspect_rows(w, h, cols, render_mode=self._get_render_mode()))
        else:
            self.rows_var.set(rows)
        self._updating_aspect = False
        self._update_res_summary()
        self._schedule_convert()

    def _on_aspect_toggle(self) -> None:
        if self.aspect_lock_var.get():
            self._apply_aspect_rows()
            self._schedule_convert()

    def _apply_aspect_rows(self) -> None:
        if self.media is None:
            return
        try:
            cols = int(self.cols_var.get())
        except (tk.TclError, ValueError):
            return
        w, h = self.media.original_size
        new_rows = max(
            10,
            min(150, compute_aspect_rows(w, h, cols, render_mode=self._get_render_mode())),
        )
        self._updating_aspect = True
        self.rows_var.set(new_rows)
        self._updating_aspect = False
        self._update_res_summary()

    def _on_cols_changed(self, *_args) -> None:
        if self._updating_aspect:
            return
        if self.aspect_lock_var.get() and self.media is not None:
            self._apply_aspect_rows()
        else:
            self._update_res_summary()
        self._schedule_convert()

    def _on_rows_changed(self, *_args) -> None:
        if self._updating_aspect:
            return
        self._update_res_summary()
        self._schedule_convert()

    def _update_res_summary(self) -> None:
        try:
            cols = max(1, int(self.cols_var.get()))
            rows = max(1, int(self.rows_var.get()))
        except (tk.TclError, ValueError):
            return
        if self._get_render_mode() == "color_cell":
            self.res_summary_var.set(f"픽셀: {cols} × {rows} px ({cols * rows:,} 셀)")
        else:
            self.res_summary_var.set(f"도트: {cols * 2} × {rows * 4} px ({cols * rows:,} 셀)")

    def _on_threshold_slide(self, val: str) -> None:
        v = int(float(val))
        self.threshold_label_var.set(str(v))
        self._schedule_convert()

    def _on_contrast_slide(self, val: str) -> None:
        self.contrast_label_var.set(f"{float(val):.1f}")
        self._schedule_convert()

    def _on_sharpness_slide(self, val: str) -> None:
        self.sharpness_label_var.set(f"{float(val):.1f}")
        self._schedule_convert()

    def _on_theme_selected(self, _event=None) -> None:
        pair = THEME_PRESETS.get(self.theme_var.get())
        if pair:
            bg, fg = pair
            self.bg_hex_var.set(bg)
            self.fg_hex_var.set(fg)
            self.bg_swatch.configure(bg=bg)
            self.fg_swatch.configure(bg=fg)
            if self._get_render_mode() == "color_cell":
                self._schedule_convert()
            else:
                self._render_current_frame()

    def _pick_bg_color(self) -> None:
        _, hex_c = colorchooser.askcolor(color=self.bg_hex_var.get(), title="배경색 선택")
        if hex_c:
            self.bg_hex_var.set(hex_c.upper())
            self.bg_swatch.configure(bg=hex_c)
            if self._get_render_mode() == "color_cell":
                self._schedule_convert()
            else:
                self._render_current_frame()

    def _pick_fg_color(self) -> None:
        _, hex_c = colorchooser.askcolor(color=self.fg_hex_var.get(), title="도트색 선택")
        if hex_c:
            self.fg_hex_var.set(hex_c.upper())
            self.fg_swatch.configure(bg=hex_c)
            self._render_current_frame()

    def _schedule_convert(self) -> None:
        if self.media is None:
            return
        if self._convert_after_id is not None:
            self.root.after_cancel(self._convert_after_id)
        self._convert_after_id = self.root.after(120, self._run_convert)

    def _read_config(self) -> ConversionConfig:
        cols = max(10, min(200, int(self.cols_var.get() or 80)))
        rows = max(10, min(150, int(self.rows_var.get() or 45)))
        dither: DitherMode = (
            "floyd_steinberg" if self.dither_var.get() == "floyd_steinberg" else "threshold"
        )
        threshold = max(1, min(254, int(self.threshold_var.get())))
        frame_step = max(1, min(10, int(self.frame_step_var.get() or 1)))
        max_frames = max(2, min(300, int(self.max_frames_var.get() or 60)))
        palette_colors = max(2, min(32, int(self.palette_colors_var.get() or 32)))
        return ConversionConfig(
            cols=cols,
            rows=rows,
            render_mode=self._get_render_mode(),
            dither=dither,
            threshold=threshold,
            invert=bool(self.invert_var.get()),
            auto_contrast=bool(self.auto_contrast_var.get()),
            contrast=float(self.contrast_var.get()),
            sharpness=float(self.sharpness_var.get()),
            frame_step=frame_step,
            max_frames=max_frames,
            palette_colors=palette_colors,
            bg_hex=self.bg_hex_var.get(),
            fg_hex=self.fg_hex_var.get(),
        )

    def _run_convert(self) -> None:
        self._convert_after_id = None
        if self.media is None:
            return
        try:
            cfg = self._read_config()
        except (tk.TclError, ValueError):
            return

        self.converted = convert_media(self.media, cfg)
        n = len(self.converted.durations_ms)
        if self.current_frame_idx >= n:
            self.current_frame_idx = 0

        self.frame_scale.configure(to=max(0, n - 1))
        self.frame_pos_var.set(self.current_frame_idx)
        self.frame_text_var.set(f"{self.current_frame_idx + 1} / {n}")

        if cfg.render_mode == "color_braille" and n > 1:
            mode_hint = "컬러 점자는 첫 프레임만 엑셀로 저장됩니다 (GIF는 컬러 셀/단색 점자 권장)"
        elif cfg.render_mode == "color_cell" and n > 1:
            mode_hint = f"GIF {n}프레임 ({cfg.palette_colors}색 팔레트, F9로 재생)"
        elif n > 1:
            mode_hint = f"GIF {n}프레임 준비됨 (단색 점자 고속, F9로 재생)"
        else:
            mode_hint = "정적 이미지 준비됨 (수식 없이 즉시 열림)"
        self.status_var.set(mode_hint)
        self._render_current_frame()

    def _render_current_frame(self) -> None:
        if self.converted is None or not self.converted.durations_ms:
            return

        n = len(self.converted.durations_ms)
        idx = max(0, min(self.current_frame_idx, n - 1))
        if self.show_original_var.get() and self.media is not None:
            step = max(1, int(self.frame_step_var.get() or 1))
            orig_idx = min(idx * step, len(self.media.frames) - 1)
            pil_img = self.media.frames[orig_idx].convert("RGB")
        else:
            pil_img = render_preview_image(
                self.converted,
                idx,
                bg_hex=self.bg_hex_var.get(),
                fg_hex=self.fg_hex_var.get(),
            )

        cw = max(100, self.preview_canvas.winfo_width())
        ch = max(100, self.preview_canvas.winfo_height())
        iw, ih = pil_img.size
        scale = min((cw - 20) / iw, (ch - 20) / ih)
        new_w = max(1, int(iw * scale))
        new_h = max(1, int(ih * scale))

        resample = (
            Image.Resampling.BILINEAR
            if self.show_original_var.get()
            else Image.Resampling.NEAREST
        )
        resized = pil_img.resize((new_w, new_h), resample)
        self._preview_photo = ImageTk.PhotoImage(resized)

        self.preview_canvas.configure(bg=self.bg_hex_var.get())
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(cw // 2, ch // 2, image=self._preview_photo, anchor=tk.CENTER)

    def _toggle_play(self) -> None:
        if self.is_playing:
            self._stop_playback()
        else:
            self._start_playback()

    def _start_playback(self) -> None:
        if self.converted is None or len(self.converted.durations_ms) <= 1:
            return
        self.is_playing = True
        self.play_btn.configure(text="⏸ 일시정지")
        self._tick_playback()

    def _stop_playback(self) -> None:
        self.is_playing = False
        self.play_btn.configure(text="▶ 재생")
        if self._play_after_id is not None:
            self.root.after_cancel(self._play_after_id)
            self._play_after_id = None

    def _tick_playback(self) -> None:
        if not self.is_playing or self.converted is None:
            return
        n = len(self.converted.durations_ms)
        if n <= 1:
            return
        self.current_frame_idx = (self.current_frame_idx + 1) % n
        self.frame_pos_var.set(self.current_frame_idx)
        self.frame_text_var.set(f"{self.current_frame_idx + 1} / {n}")
        self._render_current_frame()

        delay = self.converted.durations_ms[self.current_frame_idx]
        self._play_after_id = self.root.after(max(30, delay), self._tick_playback)

    def _on_frame_seek(self, val: str) -> None:
        if self.converted is None:
            return
        n = len(self.converted.durations_ms)
        if n == 0:
            return
        idx = max(0, min(int(float(val)), n - 1))
        self.current_frame_idx = idx
        self.frame_text_var.set(f"{idx + 1} / {n}")
        self._render_current_frame()

    def _on_export(self) -> None:
        if self.converted is None or self.media is None:
            return

        is_anim = (
            len(self.converted.durations_ms) > 1
            and self.converted.render_mode != "color_braille"
        )
        if is_anim:
            ext = ".xlsm"
            ftypes = [
                ("Excel 매크로 통합 문서 (자동 재생)", "*.xlsm"),
                ("Excel 통합 문서 (F9 수동 재생)", "*.xlsx"),
            ]
        else:
            ext = ".xlsx"
            ftypes = [("Excel 통합 문서", "*.xlsx")]

        default_name = f"{self.media.path.stem}_pixelart{ext}"
        out_path = filedialog.asksaveasfilename(
            title="엑셀 파일 저장",
            defaultextension=ext,
            initialfile=default_name,
            filetypes=ftypes,
        )
        if not out_path:
            return

        style = ExportStyle(
            bg_hex=self.bg_hex_var.get(),
            fg_hex=self.fg_hex_var.get(),
        )
        self.export_btn.configure(state=tk.DISABLED)
        self.progress["value"] = 0
        self.status_var.set("엑셀 파일 생성 중...")

        result_snapshot = self.converted

        def worker() -> None:
            try:
                def on_prog(done: int, total: int) -> None:
                    pct = int((done / max(1, total)) * 100)
                    self.root.after(0, lambda p=pct: self.progress.configure(value=p))

                export_to_xlsx(result_snapshot, out_path, style=style, progress_cb=on_prog)
                self.root.after(0, lambda: self._on_export_done(out_path, None))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_export_done(out_path, e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_export_done(self, out_path: str, err: Exception | None) -> None:
        self.export_btn.configure(state=tk.NORMAL)
        if err is not None:
            self.progress["value"] = 0
            self.status_var.set("내보내기 실패")
            messagebox.showerror("내보내기 오류", f"엑셀 생성 중 오류가 발생했습니다:\n{err}")
            return

        self.progress["value"] = 100
        self.status_var.set(f"저장 완료: {Path(out_path).name}")
        is_anim = (
            self.converted is not None
            and len(self.converted.durations_ms) > 1
            and self.converted.render_mode != "color_braille"
        )
        if is_anim and Path(out_path).suffix.lower() == ".xlsm":
            tip = "\n\nGIF 자동 재생: 엑셀에서 연 뒤 상단의 [콘텐츠 사용]을 누르면 자동 재생됩니다."
        elif is_anim:
            tip = "\n\nGIF 수동 재생: 엑셀에서 연 뒤 [F9] 키를 누르면 프레임이 재생됩니다."
        else:
            tip = ""
        messagebox.showinfo("완료", f"엑셀 파일이 저장되었습니다:\n{out_path}{tip}")


def launch_app() -> None:
    """Start the DotXLSX Tkinter application."""
    root = tk.Tk()
    DotXLSXApp(root)
    root.mainloop()
