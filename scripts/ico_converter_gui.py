#!/usr/bin/env python3
"""
ICO Converter GUI (updated)

A Windows-friendly graphical batch converter for transparent PNG/SVG artwork to .ico files.
Designed for folder icons, app shortcuts, and general Windows icon use.

Required:
    py -3 -m pip install pillow

Optional SVG support:
    Install Inkscape and leave SVG renderer set to Auto or Inkscape.
    CairoSVG can also work, but on Windows it requires the native Cairo/GTK libraries.
"""

from __future__ import annotations

import csv
import os
import queue
import re
import shutil
import subprocess
import tempfile
import sys
import threading
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Iterable, Iterator

import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

try:
    from PIL import Image, ImageColor, ImageOps, UnidentifiedImageError
except ImportError:  # pragma: no cover
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing Dependency",
        "Pillow is required.\n\nInstall it with:\n\npy -3 -m pip install pillow",
    )
    raise SystemExit(1)

try:
    RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE_LANCZOS = Image.LANCZOS

SUPPORTED_EXTENSIONS = {".png", ".svg"}
DEFAULT_ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
COMMON_PADDING_PERCENTS = [0, 2, 4, 5, 6, 8, 10, 12, 15, 20]
APP_TITLE = "ICO Batch Converter"
ACCENT = "#1976d2"
SUCCESS = "#2e7d32"
WARNING = "#ef6c00"
SOFT_BG = "#f5f8fc"


@dataclass
class ConversionResult:
    source: Path
    target: Path | None
    status: str
    message: str
    source_size: str = ""
    ico_sizes: str = ""


@dataclass
class ConversionOptions:
    output_dir: Path | None
    recursive: bool
    preserve_tree: bool
    sizes: list[int]
    fit: str
    padding: float
    background: tuple[int, int, int, int] | None
    supersample: int
    svg_renderer: str
    existing_mode: str
    dry_run: bool


@dataclass
class WorkerSettings:
    inputs: list[str]
    options: ConversionOptions
    auto_report: bool
    report_path: Path | None
    write_desktop_ini: bool
    apply_windows_attributes: bool


class ToolTip:
    def __init__(self, widget: tk.Widget, text: str, delay: int = 450) -> None:
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tipwindow: tk.Toplevel | None = None
        self.after_id: str | None = None
        self.widget.bind("<Enter>", self.schedule_show, add="+")
        self.widget.bind("<Leave>", self.hide_tip, add="+")
        self.widget.bind("<ButtonPress>", self.hide_tip, add="+")

    def schedule_show(self, _event=None) -> None:
        self.cancel_scheduled()
        self.after_id = self.widget.after(self.delay, self.show_tip)

    def cancel_scheduled(self) -> None:
        if self.after_id is not None:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def show_tip(self) -> None:
        self.after_id = None
        if self.tipwindow or not self.text:
            return
        try:
            x, y, width, height = self.widget.bbox("insert")
        except Exception:
            x = y = width = height = 0
        x = x + self.widget.winfo_rootx() + 16
        y = y + self.widget.winfo_rooty() + height + 16
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background="#fff8dc",
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=5,
            wraplength=320,
            font=("Segoe UI", 9),
        )
        label.pack()

    def hide_tip(self, _event=None) -> None:
        self.cancel_scheduled()
        if self.tipwindow is not None:
            self.tipwindow.destroy()
            self.tipwindow = None


def parse_background(value: str) -> tuple[int, int, int, int] | None:
    normalized = value.strip().lower()
    if normalized in {"", "none", "transparent", "alpha"}:
        return None
    try:
        return ImageColor.getcolor(value, "RGBA")
    except ValueError as exc:
        raise ValueError(
            "Background must be Transparent, a named color, or a HEX value like #ffffff."
        ) from exc


def iter_source_files(inputs: Iterable[str], recursive: bool) -> Iterator[tuple[Path, Path]]:
    for raw_input in inputs:
        input_path = Path(raw_input).expanduser().resolve()

        if not input_path.exists():
            yield input_path, input_path.parent
            continue

        if input_path.is_file():
            if input_path.suffix.lower() in SUPPORTED_EXTENSIONS:
                yield input_path, input_path.parent
            continue

        if input_path.is_dir():
            candidates = input_path.rglob("*") if recursive else input_path.glob("*")
            for candidate in candidates:
                if candidate.is_file() and candidate.suffix.lower() in SUPPORTED_EXTENSIONS:
                    yield candidate.resolve(), input_path


def build_target_path(source: Path, root: Path, output_dir: Path | None, preserve_tree: bool) -> Path:
    if output_dir is None:
        target_dir = source.parent
    else:
        target_dir = output_dir.expanduser().resolve()
        if preserve_tree:
            try:
                target_dir = target_dir / source.parent.relative_to(root)
            except ValueError:
                pass
    return target_dir / f"{source.stem}.ico"


def uniquify_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 10000):
        candidate = path.with_name(f"{path.stem} ({index}){path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Could not create a unique file name for {path}")


def is_missing_cairo_error(exc: BaseException) -> bool:
    text = str(exc).lower()
    return any(
        marker in text
        for marker in [
            "no library called",
            "cairo-2",
            "libcairo",
            "cannot load library",
        ]
    )


def concise_error(exc: BaseException, max_length: int = 260) -> str:
    text = " ".join(str(exc).split())
    if len(text) > max_length:
        return text[: max_length - 3] + "..."
    return text


def find_inkscape() -> str | None:
    exe = shutil.which("inkscape") or shutil.which("inkscape.exe")
    if exe:
        return exe

    possible_roots = [
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
        os.environ.get("LOCALAPPDATA"),
    ]
    possible_paths: list[Path] = []
    for root in possible_roots:
        if root:
            possible_paths.extend(
                [
                    Path(root) / "Inkscape" / "bin" / "inkscape.exe",
                    Path(root) / "Programs" / "Inkscape" / "bin" / "inkscape.exe",
                ]
            )

    for path in possible_paths:
        if path.exists():
            return str(path)
    return None


def render_svg_with_cairosvg(path: Path, render_width: int) -> Image.Image:
    try:
        import cairosvg  # type: ignore
    except ImportError as exc:
        raise RuntimeError("CairoSVG is not installed. Install it with: py -3 -m pip install cairosvg") from exc

    try:
        png_bytes = cairosvg.svg2png(url=str(path), output_width=render_width)
    except Exception as exc:  # noqa: BLE001
        if is_missing_cairo_error(exc):
            raise RuntimeError(
                "CairoSVG is installed, but the native Cairo graphics library was not found. "
                "Install GTK/Cairo for Windows or use the Inkscape SVG renderer."
            ) from exc
        raise RuntimeError(f"CairoSVG could not render this SVG: {concise_error(exc)}") from exc

    image = Image.open(BytesIO(png_bytes))
    image.load()
    return image.convert("RGBA")


def render_svg_with_inkscape(path: Path, render_width: int) -> Image.Image:
    exe = find_inkscape()
    if not exe:
        raise RuntimeError(
            "Inkscape was not found. Install Inkscape or add inkscape.exe to PATH, "
            "then choose the Inkscape or Auto SVG renderer."
        )

    with tempfile.TemporaryDirectory(prefix="ico-svg-render-") as temp_dir:
        output_png = Path(temp_dir) / "rendered.png"
        command = [
            exe,
            str(path),
            "--export-type=png",
            f"--export-filename={output_png}",
            f"--export-width={render_width}",
            "--export-background-opacity=0",
        ]
        completed = subprocess.run(command, capture_output=True, text=True, check=False)
        if completed.returncode != 0 or not output_png.exists():
            stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown Inkscape export error."
            raise RuntimeError(f"Inkscape could not render this SVG: {concise_error(RuntimeError(stderr))}")

        image = Image.open(output_png)
        image.load()
        return image.convert("RGBA")


def load_svg(path: Path, render_width: int, renderer: str) -> Image.Image:
    renderer = renderer.lower().strip()
    errors: list[str] = []

    if renderer in {"auto", "cairosvg"}:
        try:
            return render_svg_with_cairosvg(path, render_width)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"CairoSVG: {concise_error(exc)}")
            if renderer == "cairosvg":
                raise RuntimeError(errors[-1]) from exc

    if renderer in {"auto", "inkscape"}:
        try:
            return render_svg_with_inkscape(path, render_width)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Inkscape: {concise_error(exc)}")
            if renderer == "inkscape":
                raise RuntimeError(errors[-1]) from exc

    if not errors:
        errors.append(f"Unknown SVG renderer: {renderer}")

    raise RuntimeError(
        "Could not render SVG. "
        + " | ".join(errors)
        + " Recommended fix: install Inkscape, then use SVG renderer = Auto or Inkscape."
    )


def load_image(path: Path, max_icon_size: int, supersample: int, svg_renderer: str) -> Image.Image:
    if not path.exists():
        raise RuntimeError("Source file does not exist.")

    if path.suffix.lower() == ".svg":
        render_width = max_icon_size * max(1, supersample)
        return load_svg(path, render_width=render_width, renderer=svg_renderer)

    try:
        image = Image.open(path)
        image.load()
    except UnidentifiedImageError as exc:
        raise RuntimeError("Could not identify image file.") from exc
    return image.convert("RGBA")


def prepare_square_icon(
    image: Image.Image,
    size: int,
    fit: str,
    padding: float,
    background: tuple[int, int, int, int] | None,
) -> Image.Image:
    image = image.convert("RGBA")
    canvas_color = (0, 0, 0, 0) if background is None else background
    canvas = Image.new("RGBA", (size, size), canvas_color)
    inner_size = max(1, int(round(size * (1 - (padding * 2)))))

    if fit == "stretch":
        prepared = image.resize((inner_size, inner_size), RESAMPLE_LANCZOS)
    elif fit == "cover":
        prepared = ImageOps.fit(
            image,
            (inner_size, inner_size),
            method=RESAMPLE_LANCZOS,
            centering=(0.5, 0.5),
        )
    else:
        prepared = image.copy()
        prepared.thumbnail((inner_size, inner_size), RESAMPLE_LANCZOS)

    left = (size - prepared.width) // 2
    top = (size - prepared.height) // 2
    canvas.alpha_composite(prepared, (left, top))
    return canvas


def convert_one(source: Path, root: Path, options: ConversionOptions) -> ConversionResult:
    target = build_target_path(source, root, options.output_dir, options.preserve_tree)
    ico_sizes = ",".join(str(size) for size in options.sizes)

    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return ConversionResult(source, None, "skipped", "Unsupported file type.", ico_sizes=ico_sizes)

    if target.exists() and options.existing_mode == "skip":
        return ConversionResult(source, target, "skipped", "Target already exists.", ico_sizes=ico_sizes)

    if target.exists() and options.existing_mode == "unique":
        target = uniquify_path(target)

    if options.dry_run:
        return ConversionResult(source, target, "dry-run", "Would convert.", ico_sizes=ico_sizes)

    try:
        image = load_image(
            source,
            max_icon_size=max(options.sizes),
            supersample=options.supersample,
            svg_renderer=options.svg_renderer,
        )
        source_size = f"{image.width}x{image.height}"
        base_icon = prepare_square_icon(
            image=image,
            size=max(options.sizes),
            fit=options.fit,
            padding=options.padding,
            background=options.background,
        )

        target.parent.mkdir(parents=True, exist_ok=True)
        base_icon.save(target, format="ICO", sizes=[(size, size) for size in options.sizes])

        return ConversionResult(source, target, "converted", "OK", source_size, ico_sizes)
    except Exception as exc:  # noqa: BLE001
        return ConversionResult(source, target, "error", str(exc), ico_sizes=ico_sizes)


def write_report(results: list[ConversionResult], report_path: Path) -> None:
    report_path = report_path.expanduser().resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", newline="", encoding="utf-8-sig") as report_file:
        writer = csv.DictWriter(
            report_file,
            fieldnames=["status", "source", "target", "source_size", "ico_sizes", "message"],
        )
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    "status": result.status,
                    "source": str(result.source),
                    "target": str(result.target or ""),
                    "source_size": result.source_size,
                    "ico_sizes": result.ico_sizes,
                    "message": result.message,
                }
            )


def write_desktop_ini(folder: Path, icon_path: Path, apply_attributes: bool) -> str:
    ini_path = folder / "desktop.ini"
    try:
        icon_reference = icon_path.relative_to(folder)
    except ValueError:
        icon_reference = icon_path

    ini_text = (
        "[.ShellClassInfo]\n"
        f"IconResource={icon_reference},0\n"
        f"IconFile={icon_reference}\n"
        "IconIndex=0\n"
    )
    ini_path.write_text(ini_text, encoding="utf-8")

    if apply_attributes and os.name == "nt":
        subprocess.run(["attrib", "+s", str(folder)], check=False, capture_output=True)
        subprocess.run(["attrib", "+h", "+s", str(ini_path)], check=False, capture_output=True)
        if icon_path.exists():
            subprocess.run(["attrib", "+h", str(icon_path)], check=False, capture_output=True)

    return f"Wrote {ini_path}"


def write_desktop_ini_for_results(results: list[ConversionResult], apply_attributes: bool) -> list[str]:
    generated_by_folder: dict[Path, list[Path]] = defaultdict(list)
    for result in results:
        if result.status == "converted" and result.target is not None:
            generated_by_folder[result.target.parent].append(result.target)

    messages: list[str] = []
    for folder, icons in sorted(generated_by_folder.items()):
        if len(icons) != 1:
            messages.append(
                f"Skipped desktop.ini for {folder}: found {len(icons)} generated icons in that folder."
            )
            continue
        try:
            messages.append(write_desktop_ini(folder, icons[0], apply_attributes))
        except Exception as exc:  # noqa: BLE001
            messages.append(f"Could not write desktop.ini for {folder}: {exc}")
    return messages


class SizesMenu(ttk.Menubutton):
    def __init__(self, master, sizes: list[int], default_selected: list[int], **kwargs):
        self.display_var = tk.StringVar()
        super().__init__(master, textvariable=self.display_var, direction="below", **kwargs)
        self._sizes = sizes
        self._vars: dict[int, tk.BooleanVar] = {}
        self.menu = tk.Menu(self, tearoff=False)
        self.configure(menu=self.menu)
        for size in sizes:
            var = tk.BooleanVar(value=size in default_selected)
            self._vars[size] = var
            self.menu.add_checkbutton(label=f"{size} × {size}", variable=var, command=self.refresh_text)
        self.menu.add_separator()
        self.menu.add_command(label="Select All", command=self.select_all)
        self.menu.add_command(label="Clear All", command=self.clear_all)
        self.refresh_text()

    def selected_sizes(self) -> list[int]:
        return [size for size in self._sizes if self._vars[size].get()]

    def set_selected(self, selected: list[int]) -> None:
        selected_set = set(selected)
        for size, var in self._vars.items():
            var.set(size in selected_set)
        self.refresh_text()

    def select_all(self) -> None:
        for var in self._vars.values():
            var.set(True)
        self.refresh_text()

    def clear_all(self) -> None:
        for var in self._vars.values():
            var.set(False)
        self.refresh_text()

    def refresh_text(self) -> None:
        selected = self.selected_sizes()
        if not selected:
            self.display_var.set("Choose sizes…")
        else:
            self.display_var.set(", ".join(str(size) for size in selected))


class IcoConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("980x760")
        self.root.minsize(930, 700)
        self.root.configure(bg=SOFT_BG)

        self.work_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.source_paths: list[str] = []
        self.tooltips: list[ToolTip] = []
        self.log_visible = tk.BooleanVar(value=False)

        self.output_var = tk.StringVar(value="")
        self.use_source_folder_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.preserve_tree_var = tk.BooleanVar(value=True)
        self.fit_var = tk.StringVar(value="contain")
        self.padding_mode_var = tk.StringVar(value="preset")
        self.padding_preset_var = tk.IntVar(value=5)
        self.padding_custom_var = tk.StringVar(value="5")
        self.background_mode_var = tk.StringVar(value="transparent")
        self.background_custom_var = tk.StringVar(value="#ffffff")
        self.supersample_var = tk.IntVar(value=4)
        self.svg_renderer_var = tk.StringVar(value="auto")
        self.existing_mode_var = tk.StringVar(value="skip")
        self.auto_report_var = tk.BooleanVar(value=True)
        self.report_path_var = tk.StringVar(value="")
        self.write_desktop_ini_var = tk.BooleanVar(value=False)
        self.apply_attributes_var = tk.BooleanVar(value=True)

        self.total_files = 0
        self.completed_files = 0

        self.configure_style()
        self.build_ui()
        self.refresh_output_controls()
        self.refresh_padding_controls()
        self.refresh_background_controls()
        self.toggle_log_area(initial=True)

    def configure_style(self) -> None:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), background=SOFT_BG)
        style.configure("Subtle.TLabel", foreground="#52616b", background=SOFT_BG)
        style.configure("Card.TLabelframe", background="white")
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=ACCENT)
        style.configure("TFrame", background=SOFT_BG)
        style.configure("White.TFrame", background="white")
        style.configure("Primary.TButton", foreground="white", background=ACCENT, padding=(12, 8))
        style.map("Primary.TButton", background=[("active", "#1565c0")])
        style.configure("Success.TButton", foreground="white", background=SUCCESS, padding=(12, 8))
        style.map("Success.TButton", background=[("active", "#27662a")])
        style.configure("Soft.TButton", padding=(10, 7))
        style.configure("LogToggle.TButton", foreground="white", background=WARNING, padding=(10, 6))
        style.map("LogToggle.TButton", background=[("active", "#e65100")])
        style.configure("TCheckbutton", background="white")
        style.configure("TRadiobutton", background="white")
        style.configure("TLabel", background="white")

    def add_tip(self, widget: tk.Widget, text: str) -> None:
        self.tooltips.append(ToolTip(widget, text))

    def build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="ICO Batch Converter", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Convert PNG and SVG icon artwork into Windows .ico files with transparent-background support.",
            style="Subtle.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        main = ttk.Frame(outer, style="White.TFrame")
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(99, weight=1)

        self.build_sources_section(main)
        self.build_output_section(main)
        self.build_options_section(main)
        self.build_action_section(main)
        self.build_log_section(main)

    def build_sources_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Source Artwork", padding=10, style="Card.TLabelframe")
        frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.source_listbox = tk.Listbox(list_frame, height=7, selectmode="extended")
        self.source_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.source_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.source_listbox.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(frame)
        buttons.grid(row=0, column=1, sticky="ns", padx=(10, 0))
        btn_add_files = ttk.Button(buttons, text="Add Files", style="Soft.TButton", command=self.add_files)
        btn_add_files.pack(fill="x", pady=(0, 6))
        btn_add_folder = ttk.Button(buttons, text="Add Folder", style="Soft.TButton", command=self.add_folder)
        btn_add_folder.pack(fill="x", pady=(0, 6))
        btn_remove = ttk.Button(buttons, text="Remove", style="Soft.TButton", command=self.remove_selected_sources)
        btn_remove.pack(fill="x", pady=(0, 6))
        btn_clear = ttk.Button(buttons, text="Clear", style="Soft.TButton", command=self.clear_sources)
        btn_clear.pack(fill="x")

        self.add_tip(btn_add_files, "Add one or more PNG or SVG files directly.")
        self.add_tip(btn_add_folder, "Add a folder that contains PNG or SVG artwork.")
        self.add_tip(btn_remove, "Remove the selected source entries from the list.")
        self.add_tip(btn_clear, "Remove all source entries from the list.")

        chk_recursive = ttk.Checkbutton(frame, text="Search folders recursively", variable=self.recursive_var)
        chk_recursive.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.add_tip(chk_recursive, "If enabled, subfolders are searched too when you add a folder.")

    def build_output_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Output", padding=10, style="Card.TLabelframe")
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        chk_source = ttk.Checkbutton(
            frame,
            text="Save .ico files next to the source artwork",
            variable=self.use_source_folder_var,
            command=self.refresh_output_controls,
        )
        chk_source.grid(row=0, column=0, columnspan=3, sticky="w")
        self.add_tip(chk_source, "When enabled, each .ico file is saved in the same folder as its source file.")

        ttk.Label(frame, text="Output folder:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.output_entry = ttk.Entry(frame, textvariable=self.output_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
        self.output_button = ttk.Button(frame, text="Browse", style="Soft.TButton", command=self.choose_output_folder)
        self.output_button.grid(row=1, column=2, pady=(8, 0))
        self.add_tip(self.output_entry, "Choose a central output folder if you do not want files saved next to the source artwork.")
        self.add_tip(self.output_button, "Browse for a target output folder.")

        self.preserve_tree_check = ttk.Checkbutton(
            frame,
            text="Preserve source subfolder structure inside the output folder",
            variable=self.preserve_tree_var,
        )
        self.preserve_tree_check.grid(row=2, column=1, sticky="w", pady=(8, 0))
        self.add_tip(self.preserve_tree_check, "Keeps the same subfolder layout under the selected output folder.")

    def build_options_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Icon Options", padding=10, style="Card.TLabelframe")
        frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        for col in range(4):
            frame.columnconfigure(col, weight=1)

        # Row 0 labels
        lbl_sizes = ttk.Label(frame, text="ICO sizes")
        lbl_sizes.grid(row=0, column=0, sticky="w")
        lbl_fit = ttk.Label(frame, text="Fit mode")
        lbl_fit.grid(row=0, column=1, sticky="w")
        lbl_existing = ttk.Label(frame, text="Existing files")
        lbl_existing.grid(row=0, column=2, sticky="w")
        lbl_renderer = ttk.Label(frame, text="SVG renderer")
        lbl_renderer.grid(row=0, column=3, sticky="w")

        self.sizes_menu = SizesMenu(frame, DEFAULT_ICO_SIZES, DEFAULT_ICO_SIZES)
        self.sizes_menu.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        cmb_fit = ttk.Combobox(
            frame,
            textvariable=self.fit_var,
            values=["contain", "cover", "stretch"],
            state="readonly",
        )
        cmb_fit.grid(row=1, column=1, sticky="ew", padx=(0, 10))
        cmb_existing = ttk.Combobox(
            frame,
            textvariable=self.existing_mode_var,
            values=["skip", "overwrite", "unique"],
            state="readonly",
        )
        cmb_existing.grid(row=1, column=2, sticky="ew", padx=(0, 10))
        cmb_renderer = ttk.Combobox(
            frame,
            textvariable=self.svg_renderer_var,
            values=["auto", "inkscape", "cairosvg"],
            state="readonly",
        )
        cmb_renderer.grid(row=1, column=3, sticky="ew")

        lbl_svg = ttk.Label(frame, text="SVG supersample")
        lbl_svg.grid(row=2, column=3, sticky="w", pady=(10, 0))
        spn_svg = ttk.Spinbox(frame, from_=1, to=8, textvariable=self.supersample_var, width=8)
        spn_svg.grid(row=3, column=3, sticky="ew", pady=(0, 0))

        self.add_tip(self.sizes_menu, "Choose one or more embedded icon sizes to include in the ICO file.")
        self.add_tip(cmb_fit, "Contain = preserve the whole image. Cover = fill the icon and crop overflow. Stretch = force the image into a square.")
        self.add_tip(cmb_existing, "Skip leaves existing ICOs alone. Overwrite replaces them. Unique creates files like name (2).ico.")
        self.add_tip(cmb_renderer, "Auto tries CairoSVG first, then Inkscape. Choose Inkscape if CairoSVG produces missing Cairo DLL errors on Windows.")
        self.add_tip(spn_svg, "Higher values can improve SVG edge quality, but may slightly slow conversion.")

        # Padding section
        pad_box = ttk.LabelFrame(frame, text="Padding", padding=8, style="Card.TLabelframe")
        pad_box.grid(row=2, column=0, columnspan=2, sticky="ew", padx=(0, 10), pady=(12, 0))
        pad_box.columnconfigure(1, weight=1)
        rad_pad_preset = ttk.Radiobutton(
            pad_box,
            text="Preset",
            value="preset",
            variable=self.padding_mode_var,
            command=self.refresh_padding_controls,
        )
        rad_pad_custom = ttk.Radiobutton(
            pad_box,
            text="Custom %",
            value="custom",
            variable=self.padding_mode_var,
            command=self.refresh_padding_controls,
        )
        rad_pad_preset.grid(row=0, column=0, sticky="w")
        rad_pad_custom.grid(row=1, column=0, sticky="w", pady=(6, 0))
        cmb_padding = ttk.Combobox(
            pad_box,
            textvariable=self.padding_preset_var,
            values=COMMON_PADDING_PERCENTS,
            state="readonly",
            width=8,
        )
        cmb_padding.grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(pad_box, text="% padding around the icon artwork").grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.padding_custom_entry = ttk.Entry(pad_box, textvariable=self.padding_custom_var, width=10)
        self.padding_custom_entry.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(6, 0))
        ttk.Label(pad_box, text="Enter 0–40").grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(6, 0))
        self.add_tip(pad_box, "Padding adds empty space around the artwork. This helps icons avoid looking cramped in Windows.")
        self.add_tip(cmb_padding, "Common padding presets for fast setup.")
        self.add_tip(self.padding_custom_entry, "Enter a custom padding percent between 0 and 40.")

        # Background section
        bg_box = ttk.LabelFrame(frame, text="Background", padding=8, style="Card.TLabelframe")
        bg_box.grid(row=2, column=2, columnspan=2, sticky="ew", pady=(12, 0))
        bg_box.columnconfigure(1, weight=1)
        rad_bg_trans = ttk.Radiobutton(
            bg_box,
            text="Transparent",
            value="transparent",
            variable=self.background_mode_var,
            command=self.refresh_background_controls,
        )
        rad_bg_white = ttk.Radiobutton(
            bg_box,
            text="White",
            value="white",
            variable=self.background_mode_var,
            command=self.refresh_background_controls,
        )
        rad_bg_black = ttk.Radiobutton(
            bg_box,
            text="Black",
            value="black",
            variable=self.background_mode_var,
            command=self.refresh_background_controls,
        )
        rad_bg_custom = ttk.Radiobutton(
            bg_box,
            text="Custom",
            value="custom",
            variable=self.background_mode_var,
            command=self.refresh_background_controls,
        )
        rad_bg_trans.grid(row=0, column=0, sticky="w")
        rad_bg_white.grid(row=0, column=1, sticky="w")
        rad_bg_black.grid(row=1, column=0, sticky="w", pady=(6, 0))
        rad_bg_custom.grid(row=1, column=1, sticky="w", pady=(6, 0))
        self.background_custom_entry = ttk.Entry(bg_box, textvariable=self.background_custom_var, width=12)
        self.background_custom_entry.grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.background_pick_button = ttk.Button(bg_box, text="Pick Color", style="Soft.TButton", command=self.pick_background_color)
        self.background_pick_button.grid(row=2, column=1, sticky="w", pady=(8, 0))
        self.background_preview = tk.Label(bg_box, width=3, relief="solid", borderwidth=1, bg="#ffffff")
        self.background_preview.grid(row=2, column=2, sticky="w", padx=(8, 0), pady=(8, 0))
        self.add_tip(bg_box, "Choose whether the icon canvas should stay transparent or be filled with a background color.")
        self.add_tip(self.background_custom_entry, "For Custom, enter a HEX color such as #ffffff or click Pick Color.")
        self.add_tip(self.background_pick_button, "Choose a custom background color visually.")

        # Additional options row
        row3 = ttk.Frame(frame)
        row3.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        row3.columnconfigure(1, weight=1)
        chk_report = ttk.Checkbutton(row3, text="Create CSV report", variable=self.auto_report_var)
        chk_report.grid(row=0, column=0, sticky="w")
        chk_ini = ttk.Checkbutton(row3, text="Write desktop.ini for folder icons", variable=self.write_desktop_ini_var)
        chk_ini.grid(row=0, column=1, sticky="w", padx=(16, 0))
        chk_attr = ttk.Checkbutton(
            row3,
            text="Apply Windows folder/icon attributes",
            variable=self.apply_attributes_var,
        )
        chk_attr.grid(row=0, column=2, sticky="w", padx=(16, 0))
        self.add_tip(chk_report, "Writes a CSV report of converted, skipped, and failed files.")
        self.add_tip(chk_ini, "Creates desktop.ini for folders that contain exactly one generated icon. Useful for folder icon customization.")
        self.add_tip(chk_attr, "Applies hidden/system attributes used by Windows folder icon behavior when desktop.ini is written.")

        report_row = ttk.Frame(frame)
        report_row.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        report_row.columnconfigure(1, weight=1)
        ttk.Label(report_row, text="Optional report path").grid(row=0, column=0, sticky="w")
        report_entry = ttk.Entry(report_row, textvariable=self.report_path_var)
        report_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        report_btn = ttk.Button(report_row, text="Browse", style="Soft.TButton", command=self.choose_report_path)
        report_btn.grid(row=0, column=2)
        self.add_tip(report_entry, "Leave blank to auto-create a timestamped report when CSV reporting is enabled.")
        self.add_tip(report_btn, "Choose a specific CSV report file path.")

        for w, txt in [
            (lbl_sizes, "Choose which embedded sizes should be included in the ICO file."),
            (lbl_fit, "Controls how the source artwork is fitted onto the square icon canvas."),
            (lbl_existing, "What should happen if a target ICO file already exists."),
            (lbl_renderer, "Controls which engine is used to turn SVG files into PNG before creating the ICO."),
            (lbl_svg, "Controls SVG rasterization quality before the image is saved as ICO."),
        ]:
            self.add_tip(w, txt)

    def build_action_section(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        btn_convert = ttk.Button(
            frame,
            text="Convert Icons",
            style="Primary.TButton",
            command=lambda: self.start_conversion(dry_run=False),
        )
        btn_convert.pack(side="left")
        btn_dry = ttk.Button(
            frame,
            text="Dry Run",
            style="Success.TButton",
            command=lambda: self.start_conversion(dry_run=True),
        )
        btn_dry.pack(side="left", padx=(8, 0))
        btn_output = ttk.Button(frame, text="Open Output Folder", style="Soft.TButton", command=self.open_output_folder)
        btn_output.pack(side="left", padx=(8, 0))
        self.log_toggle_btn = ttk.Button(frame, text="Show Log", style="LogToggle.TButton", command=self.toggle_log_area)
        self.log_toggle_btn.pack(side="left", padx=(8, 0))
        btn_exit = ttk.Button(frame, text="Exit", style="Soft.TButton", command=self.root.destroy)
        btn_exit.pack(side="right")

        self.convert_button = btn_convert
        self.dry_run_button = btn_dry
        self.add_tip(btn_convert, "Convert all discovered PNG and SVG files into ICO files using the current settings.")
        self.add_tip(btn_dry, "Preview what would happen without writing any files.")
        self.add_tip(btn_output, "Open the selected output folder, or the first source folder when output is saved next to the source files.")
        self.add_tip(self.log_toggle_btn, "Show or hide the log section.")
        self.add_tip(btn_exit, "Close the application.")

        self.progress = ttk.Progressbar(parent, mode="determinate")
        self.progress.grid(row=4, column=0, sticky="ew")

    def build_log_section(self, parent: ttk.Frame) -> None:
        self.log_container = ttk.LabelFrame(parent, text="Log", padding=10, style="Card.TLabelframe")
        self.log_container.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        parent.rowconfigure(5, weight=1)
        self.log_container.columnconfigure(0, weight=1)
        self.log_container.rowconfigure(0, weight=1)

        self.log_text = tk.Text(self.log_container, wrap="word", height=10)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(self.log_container, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.add_tip(self.log_text, "Shows progress, summary information, and any conversion errors.")

    def toggle_log_area(self, initial: bool = False) -> None:
        visible = self.log_visible.get()
        if initial:
            visible = False
        else:
            visible = not visible
        self.log_visible.set(visible)
        if visible:
            self.log_container.grid()
            self.log_toggle_btn.configure(text="Hide Log")
        else:
            self.log_container.grid_remove()
            if hasattr(self, "log_toggle_btn"):
                self.log_toggle_btn.configure(text="Show Log")

    def refresh_output_controls(self) -> None:
        use_source_folder = self.use_source_folder_var.get()
        state = "disabled" if use_source_folder else "normal"
        self.output_entry.configure(state=state)
        self.output_button.configure(state=state)
        self.preserve_tree_check.configure(state=state)

    def refresh_padding_controls(self) -> None:
        custom = self.padding_mode_var.get() == "custom"
        self.padding_custom_entry.configure(state="normal" if custom else "disabled")

    def refresh_background_controls(self) -> None:
        custom = self.background_mode_var.get() == "custom"
        state = "normal" if custom else "disabled"
        self.background_custom_entry.configure(state=state)
        self.background_pick_button.configure(state=state)
        self.update_background_preview()

    def update_background_preview(self) -> None:
        mode = self.background_mode_var.get()
        preview = "#ffffff"
        if mode == "transparent":
            preview = "#ffffff"
        elif mode == "white":
            preview = "#ffffff"
        elif mode == "black":
            preview = "#000000"
        else:
            raw = self.background_custom_var.get().strip()
            if re.fullmatch(r"#?[0-9a-fA-F]{6}", raw):
                preview = raw if raw.startswith("#") else f"#{raw}"
        self.background_preview.configure(bg=preview)

    def add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select PNG or SVG files",
            filetypes=[("Icon artwork", "*.png *.svg"), ("PNG files", "*.png"), ("SVG files", "*.svg"), ("All files", "*.*")],
        )
        for file_path in files:
            self.add_source_path(file_path)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select a folder containing PNG/SVG artwork")
        if folder:
            self.add_source_path(folder)

    def add_source_path(self, path: str) -> None:
        resolved = str(Path(path).expanduser().resolve())
        if resolved not in self.source_paths:
            self.source_paths.append(resolved)
            self.source_listbox.insert("end", resolved)

    def remove_selected_sources(self) -> None:
        selected = list(self.source_listbox.curselection())
        for index in reversed(selected):
            self.source_listbox.delete(index)
            del self.source_paths[index]

    def clear_sources(self) -> None:
        self.source_paths.clear()
        self.source_listbox.delete(0, "end")

    def choose_output_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_var.set(folder)

    def choose_report_path(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save CSV report as",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.report_path_var.set(path)

    def pick_background_color(self) -> None:
        color = colorchooser.askcolor(title="Choose background color", initialcolor=self.background_custom_var.get())[1]
        if color:
            self.background_custom_var.set(color)
            self.update_background_preview()

    def open_output_folder(self) -> None:
        folder: Path | None = None
        if not self.use_source_folder_var.get() and self.output_var.get().strip():
            folder = Path(self.output_var.get()).expanduser().resolve()
        elif self.source_paths:
            first = Path(self.source_paths[0]).expanduser().resolve()
            folder = first if first.is_dir() else first.parent

        if not folder:
            messagebox.showinfo("No Folder", "Choose an output folder or add source artwork first.")
            return

        folder.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(folder)  # type: ignore[attr-defined]
        else:
            subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", str(folder)], check=False)

    def log(self, message: str) -> None:
        if not self.log_visible.get():
            self.toggle_log_area()
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def set_running_state(self, is_running: bool) -> None:
        state = "disabled" if is_running else "normal"
        self.convert_button.configure(state=state)
        self.dry_run_button.configure(state=state)

    def selected_padding_ratio(self) -> float:
        if self.padding_mode_var.get() == "preset":
            value = self.padding_preset_var.get()
        else:
            raw = self.padding_custom_var.get().strip()
            if raw == "":
                raise ValueError("Enter a padding percentage.")
            value = float(raw)
        if value < 0 or value > 40:
            raise ValueError("Padding must be between 0 and 40 percent.")
        return value / 100

    def selected_background_string(self) -> str:
        mode = self.background_mode_var.get()
        if mode == "transparent":
            return "transparent"
        if mode == "white":
            return "#ffffff"
        if mode == "black":
            return "#000000"
        raw = self.background_custom_var.get().strip()
        if not raw:
            raise ValueError("Enter a custom background color or pick one.")
        if re.fullmatch(r"#?[0-9a-fA-F]{6}", raw):
            return raw if raw.startswith("#") else f"#{raw}"
        return raw

    def collect_settings(self, dry_run: bool) -> WorkerSettings:
        if not self.source_paths:
            raise ValueError("Add at least one PNG/SVG file or source folder.")

        sizes = self.sizes_menu.selected_sizes()
        if not sizes:
            raise ValueError("Select at least one ICO size.")

        output_dir: Path | None = None
        if not self.use_source_folder_var.get():
            raw_output = self.output_var.get().strip()
            if not raw_output:
                raise ValueError("Choose an output folder or enable saving next to source artwork.")
            output_dir = Path(raw_output).expanduser().resolve()

        padding = self.selected_padding_ratio()
        background = parse_background(self.selected_background_string())

        supersample = int(self.supersample_var.get())
        if supersample < 1 or supersample > 8:
            raise ValueError("SVG supersample must be between 1 and 8.")

        svg_renderer = self.svg_renderer_var.get().strip().lower()
        if svg_renderer not in {"auto", "inkscape", "cairosvg"}:
            raise ValueError("SVG renderer must be Auto, Inkscape, or CairoSVG.")

        report_path = Path(self.report_path_var.get()).expanduser().resolve() if self.report_path_var.get().strip() else None

        options = ConversionOptions(
            output_dir=output_dir,
            recursive=self.recursive_var.get(),
            preserve_tree=self.preserve_tree_var.get(),
            sizes=sizes,
            fit=self.fit_var.get(),
            padding=padding,
            background=background,
            supersample=supersample,
            svg_renderer=svg_renderer,
            existing_mode=self.existing_mode_var.get(),
            dry_run=dry_run,
        )

        return WorkerSettings(
            inputs=list(self.source_paths),
            options=options,
            auto_report=self.auto_report_var.get(),
            report_path=report_path,
            write_desktop_ini=self.write_desktop_ini_var.get(),
            apply_windows_attributes=self.apply_attributes_var.get(),
        )

    def start_conversion(self, dry_run: bool) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Still Running", "A conversion is already running.")
            return

        try:
            settings = self.collect_settings(dry_run=dry_run)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Check Settings", str(exc))
            return

        self.log_text.delete("1.0", "end")
        self.progress.configure(value=0, maximum=1)
        self.completed_files = 0
        self.total_files = 0
        self.set_running_state(True)
        self.log("Starting dry run..." if dry_run else "Starting conversion...")

        self.worker_thread = threading.Thread(target=self.worker, args=(settings,), daemon=True)
        self.worker_thread.start()
        self.root.after(100, self.poll_worker_queue)

    def worker(self, settings: WorkerSettings) -> None:
        try:
            sources = list(iter_source_files(settings.inputs, settings.options.recursive))
            real_sources = [(source, root) for source, root in sources if source.exists()]
            missing_sources = [(source, root) for source, root in sources if not source.exists()]
            self.work_queue.put(("total", len(real_sources) + len(missing_sources)))

            results: list[ConversionResult] = []
            for source, _root in missing_sources:
                result = ConversionResult(source, None, "error", "Source does not exist.")
                results.append(result)
                self.work_queue.put(("result", result))

            if not real_sources:
                self.work_queue.put(("message", "No supported PNG or SVG files found."))

            for source, root in real_sources:
                result = convert_one(source, root, settings.options)
                results.append(result)
                self.work_queue.put(("result", result))

            if settings.write_desktop_ini and not settings.options.dry_run:
                for message in write_desktop_ini_for_results(results, settings.apply_windows_attributes):
                    self.work_queue.put(("message", message))

            if (settings.auto_report or settings.report_path) and results:
                report_path = settings.report_path
                if report_path is None:
                    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                    if settings.options.output_dir is not None:
                        report_base = settings.options.output_dir
                    else:
                        first_source = results[0].source
                        report_base = first_source.parent if first_source.parent.exists() else Path.cwd()
                    report_path = report_base / f"ico-conversion-report-{timestamp}.csv"

                write_report(results, report_path)
                self.work_queue.put(("message", f"Report written: {report_path}"))

            self.work_queue.put(("done", results))
        except Exception as exc:  # noqa: BLE001
            self.work_queue.put(("fatal", str(exc)))

    def poll_worker_queue(self) -> None:
        try:
            while True:
                event, payload = self.work_queue.get_nowait()

                if event == "total":
                    self.total_files = max(1, int(payload))
                    self.progress.configure(maximum=self.total_files, value=0)
                elif event == "result":
                    result = payload
                    assert isinstance(result, ConversionResult)
                    self.completed_files += 1
                    self.progress.configure(value=self.completed_files)
                    target = f" -> {result.target}" if result.target else ""
                    self.log(f"[{self.completed_files}/{self.total_files}] {result.status.upper()}: {result.source}{target}")
                    if result.message and result.message != "OK":
                        self.log(f"    {result.message}")
                elif event == "message":
                    self.log(str(payload))
                elif event == "fatal":
                    self.log(f"Fatal error: {payload}")
                    messagebox.showerror("Conversion Error", str(payload))
                    self.set_running_state(False)
                    return
                elif event == "done":
                    results = payload
                    assert isinstance(results, list)
                    self.show_done_summary(results)
                    self.set_running_state(False)
                    return
        except queue.Empty:
            pass

        if self.worker_thread and self.worker_thread.is_alive():
            self.root.after(100, self.poll_worker_queue)
        else:
            self.set_running_state(False)

    def show_done_summary(self, results: list[ConversionResult]) -> None:
        counts: dict[str, int] = defaultdict(int)
        for result in results:
            counts[result.status] += 1

        self.log("")
        self.log("Summary")
        self.log("-------")
        for status in ["converted", "skipped", "dry-run", "error"]:
            if counts.get(status):
                self.log(f"{status}: {counts[status]}")

        if counts.get("error"):
            messagebox.showwarning(
                "Finished with Errors",
                f"Finished with {counts['error']} error(s). Check the log or CSV report for details.",
            )
        else:
            messagebox.showinfo("Done", "Finished successfully.")


def main() -> int:
    root = tk.Tk()
    app = IcoConverterApp(root)
    app.update_background_preview()
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
