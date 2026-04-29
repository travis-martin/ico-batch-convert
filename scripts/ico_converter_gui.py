#!/usr/bin/env python3
"""
ICO Converter GUI

A Windows-friendly graphical batch converter for transparent PNG/SVG artwork to .ico files.
Designed for folder icons, app shortcuts, and general Windows icon use.

Required:
    pip install pillow

Optional SVG support:
    pip install cairosvg

Recommended Windows setup:
    1. Save this file as ico_converter_gui.py.
    2. Double-click it if Python is associated with .py files, or run:
       py -3 ico_converter_gui.py
    3. In the app, click "Create Launcher" to generate an Autorun .bat file and
       optionally a desktop shortcut.
"""

from __future__ import annotations

import argparse
import csv
import os
import queue
import subprocess
import sys
import threading
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Iterable, Iterator

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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
except AttributeError:  # Older Pillow fallback
    RESAMPLE_LANCZOS = Image.LANCZOS

SUPPORTED_EXTENSIONS = {".png", ".svg"}
DEFAULT_ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
APP_TITLE = "ICO Batch Converter"


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


def parse_sizes(value: str) -> list[int]:
    try:
        sizes = sorted({int(part.strip()) for part in value.split(",") if part.strip()})
    except ValueError as exc:
        raise ValueError("Sizes must be comma-separated integers, such as 16,32,48,256.") from exc

    if not sizes:
        raise ValueError("At least one icon size is required.")

    invalid = [size for size in sizes if size < 1 or size > 256]
    if invalid:
        raise ValueError(f"ICO sizes must be between 1 and 256 pixels. Invalid: {invalid}")

    return sizes


def parse_background(value: str) -> tuple[int, int, int, int] | None:
    normalized = value.strip().lower()
    if normalized in {"", "none", "transparent", "alpha"}:
        return None

    try:
        return ImageColor.getcolor(value, "RGBA")
    except ValueError as exc:
        raise ValueError(
            "Background must be transparent, none, a named color, or a HEX color like #ffffff."
        ) from exc


def parse_padding(value: str) -> float:
    raw = value.strip()
    try:
        if raw.endswith("%"):
            padding = float(raw[:-1]) / 100
        else:
            padding = float(raw)
    except ValueError as exc:
        raise ValueError("Padding must be a number such as 0.08 or 8%.") from exc

    if not 0 <= padding < 0.45:
        raise ValueError("Padding must be at least 0 and less than 0.45.")

    return padding


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


def load_svg(path: Path, render_width: int) -> Image.Image:
    try:
        import cairosvg  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "SVG conversion requires CairoSVG. Install it with: py -3 -m pip install cairosvg"
        ) from exc

    png_bytes = cairosvg.svg2png(url=str(path), output_width=render_width)
    image = Image.open(BytesIO(png_bytes))
    image.load()
    return image.convert("RGBA")


def load_image(path: Path, max_icon_size: int, supersample: int) -> Image.Image:
    if not path.exists():
        raise RuntimeError("Source file does not exist.")

    if path.suffix.lower() == ".svg":
        render_width = max_icon_size * max(1, supersample)
        return load_svg(path, render_width=render_width)

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
        image = load_image(source, max_icon_size=max(options.sizes), supersample=options.supersample)
        source_size = f"{image.width}x{image.height}"
        base_icon = prepare_square_icon(
            image=image,
            size=max(options.sizes),
            fit=options.fit,
            padding=options.padding,
            background=options.background,
        )

        target.parent.mkdir(parents=True, exist_ok=True)
        base_icon.save(
            target,
            format="ICO",
            sizes=[(size, size) for size in options.sizes],
        )

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


def get_desktop_path() -> Path:
    user_profile = os.environ.get("USERPROFILE")
    if user_profile:
        desktop = Path(user_profile) / "Desktop"
        if desktop.exists():
            return desktop
    return Path.home() / "Desktop"


def create_launcher_files(script_path: Path, launcher_dir: Path, create_desktop_shortcut: bool) -> tuple[Path, Path | None]:
    script_path = script_path.resolve()
    launcher_dir = launcher_dir.expanduser().resolve()
    launcher_dir.mkdir(parents=True, exist_ok=True)

    launcher_path = launcher_dir / "Autorun - ICO Converter.bat"
    bat_text = f"""@echo off
setlocal
cd /d "{script_path.parent}"

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    py -3 "{script_path}"
) else (
    python "{script_path}"
)

endlocal
"""
    launcher_path.write_text(bat_text, encoding="utf-8")

    shortcut_path: Path | None = None
    if create_desktop_shortcut and os.name == "nt":
        shortcut_path = get_desktop_path() / "ICO Batch Converter.lnk"
        ps_script = f"""
$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut('{str(shortcut_path).replace("'", "''")}')
$Shortcut.TargetPath = '{str(launcher_path).replace("'", "''")}';
$Shortcut.WorkingDirectory = '{str(script_path.parent).replace("'", "''")}';
$Shortcut.Description = 'Launch ICO Batch Converter';
$Shortcut.Save()
"""
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            check=False,
            capture_output=True,
            text=True,
        )

    return launcher_path, shortcut_path


class IcoConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("960x760")
        self.root.minsize(880, 680)

        self.work_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.source_paths: list[str] = []

        self.output_var = tk.StringVar(value="")
        self.use_source_folder_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.preserve_tree_var = tk.BooleanVar(value=True)
        self.sizes_var = tk.StringVar(value=",".join(str(size) for size in DEFAULT_ICO_SIZES))
        self.fit_var = tk.StringVar(value="contain")
        self.padding_var = tk.StringVar(value="5%")
        self.background_var = tk.StringVar(value="transparent")
        self.supersample_var = tk.IntVar(value=4)
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

    def configure_style(self) -> None:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Subtle.TLabel", foreground="#555555")
        style.configure("Accent.TButton", padding=(12, 7))
        style.configure("Danger.TButton", padding=(10, 6))

    def build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 12))
        ttk.Label(header, text="ICO Batch Converter", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Batch convert transparent PNG/SVG artwork into Windows .ico files for folder icons and shortcuts.",
            style="Subtle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        main_pane = ttk.PanedWindow(outer, orient="vertical")
        main_pane.pack(fill="both", expand=True)

        top_frame = ttk.Frame(main_pane)
        bottom_frame = ttk.Frame(main_pane)
        main_pane.add(top_frame, weight=3)
        main_pane.add(bottom_frame, weight=2)

        self.build_sources_section(top_frame)
        self.build_output_section(top_frame)
        self.build_options_section(top_frame)
        self.build_action_section(top_frame)
        self.build_log_section(bottom_frame)

    def build_sources_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Source Artwork", padding=10)
        frame.pack(fill="both", expand=True, pady=(0, 10))

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.source_listbox = tk.Listbox(list_frame, height=8, selectmode="extended")
        self.source_listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.source_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.source_listbox.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(frame)
        buttons.grid(row=0, column=1, sticky="ns", padx=(10, 0))

        ttk.Button(buttons, text="Add Files", command=self.add_files).pack(fill="x", pady=(0, 6))
        ttk.Button(buttons, text="Add Folder", command=self.add_folder).pack(fill="x", pady=(0, 6))
        ttk.Button(buttons, text="Remove", command=self.remove_selected_sources).pack(fill="x", pady=(0, 6))
        ttk.Button(buttons, text="Clear", command=self.clear_sources).pack(fill="x")

        ttk.Checkbutton(frame, text="Search folders recursively", variable=self.recursive_var).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )

    def build_output_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Output", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            frame,
            text="Save .ico files next to source artwork",
            variable=self.use_source_folder_var,
            command=self.refresh_output_controls,
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(frame, text="Output folder:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.output_entry = ttk.Entry(frame, textvariable=self.output_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
        self.output_button = ttk.Button(frame, text="Browse", command=self.choose_output_folder)
        self.output_button.grid(row=1, column=2, pady=(8, 0))

        self.preserve_tree_check = ttk.Checkbutton(
            frame,
            text="Preserve source subfolder structure inside output folder",
            variable=self.preserve_tree_var,
        )
        self.preserve_tree_check.grid(row=2, column=1, sticky="w", pady=(8, 0))

    def build_options_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Icon Options", padding=10)
        frame.pack(fill="x", pady=(0, 10))

        for col in range(4):
            frame.columnconfigure(col, weight=1)

        ttk.Label(frame, text="ICO sizes:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.sizes_var).grid(row=1, column=0, sticky="ew", padx=(0, 10))

        ttk.Label(frame, text="Fit mode:").grid(row=0, column=1, sticky="w")
        ttk.Combobox(
            frame,
            textvariable=self.fit_var,
            values=["contain", "cover", "stretch"],
            state="readonly",
        ).grid(row=1, column=1, sticky="ew", padx=(0, 10))

        ttk.Label(frame, text="Padding:").grid(row=0, column=2, sticky="w")
        ttk.Entry(frame, textvariable=self.padding_var).grid(row=1, column=2, sticky="ew", padx=(0, 10))

        ttk.Label(frame, text="Background:").grid(row=0, column=3, sticky="w")
        ttk.Entry(frame, textvariable=self.background_var).grid(row=1, column=3, sticky="ew")

        ttk.Label(frame, text="Existing files:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Combobox(
            frame,
            textvariable=self.existing_mode_var,
            values=["skip", "overwrite", "unique"],
            state="readonly",
        ).grid(row=3, column=0, sticky="ew", padx=(0, 10))

        ttk.Label(frame, text="SVG supersample:").grid(row=2, column=1, sticky="w", pady=(10, 0))
        ttk.Spinbox(
            frame,
            from_=1,
            to=8,
            textvariable=self.supersample_var,
            width=8,
        ).grid(row=3, column=1, sticky="ew", padx=(0, 10))

        ttk.Checkbutton(frame, text="Create CSV report", variable=self.auto_report_var).grid(
            row=3, column=2, sticky="w", padx=(0, 10)
        )

        ttk.Checkbutton(frame, text="Write desktop.ini for folder icons", variable=self.write_desktop_ini_var).grid(
            row=3, column=3, sticky="w"
        )

        ttk.Checkbutton(
            frame,
            text="Apply Windows folder/icon attributes when writing desktop.ini",
            variable=self.apply_attributes_var,
        ).grid(row=4, column=0, columnspan=4, sticky="w", pady=(10, 0))

        report_frame = ttk.Frame(frame)
        report_frame.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        report_frame.columnconfigure(1, weight=1)
        ttk.Label(report_frame, text="Optional report path:").grid(row=0, column=0, sticky="w")
        ttk.Entry(report_frame, textvariable=self.report_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(report_frame, text="Browse", command=self.choose_report_path).grid(row=0, column=2)

    def build_action_section(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(0, 10))

        self.convert_button = ttk.Button(
            frame,
            text="Convert Icons",
            style="Accent.TButton",
            command=lambda: self.start_conversion(dry_run=False),
        )
        self.convert_button.pack(side="left")

        self.dry_run_button = ttk.Button(
            frame,
            text="Dry Run",
            command=lambda: self.start_conversion(dry_run=True),
        )
        self.dry_run_button.pack(side="left", padx=(8, 0))

        ttk.Button(frame, text="Create Launcher", command=self.create_launcher).pack(side="left", padx=(8, 0))
        ttk.Button(frame, text="Open Output Folder", command=self.open_output_folder).pack(side="left", padx=(8, 0))
        ttk.Button(frame, text="Exit", command=self.root.destroy).pack(side="right")

        self.progress = ttk.Progressbar(parent, mode="determinate")
        self.progress.pack(fill="x")

    def build_log_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Log", padding=10)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(frame, wrap="word", height=10)
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def refresh_output_controls(self) -> None:
        use_source_folder = self.use_source_folder_var.get()
        state = "disabled" if use_source_folder else "normal"
        self.output_entry.configure(state=state)
        self.output_button.configure(state=state)
        self.preserve_tree_check.configure(state=state)

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
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def set_running_state(self, is_running: bool) -> None:
        state = "disabled" if is_running else "normal"
        self.convert_button.configure(state=state)
        self.dry_run_button.configure(state=state)

    def collect_settings(self, dry_run: bool) -> WorkerSettings:
        if not self.source_paths:
            raise ValueError("Add at least one PNG/SVG file or source folder.")

        output_dir: Path | None = None
        if not self.use_source_folder_var.get():
            raw_output = self.output_var.get().strip()
            if not raw_output:
                raise ValueError("Choose an output folder or enable saving next to source artwork.")
            output_dir = Path(raw_output).expanduser().resolve()

        sizes = parse_sizes(self.sizes_var.get())
        padding = parse_padding(self.padding_var.get())
        background = parse_background(self.background_var.get())
        supersample = int(self.supersample_var.get())
        if supersample < 1:
            raise ValueError("SVG supersample must be 1 or greater.")

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
            for source, root in missing_sources:
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

    def create_launcher(self) -> None:
        script_path = Path(__file__).resolve()
        launcher_dir = filedialog.askdirectory(
            title="Choose where to save the Autorun launcher",
            initialdir=str(script_path.parent),
        )
        if not launcher_dir:
            return

        create_shortcut = messagebox.askyesno(
            "Desktop Shortcut",
            "Create a desktop shortcut that points to the Autorun launcher?",
        )

        try:
            launcher_path, shortcut_path = create_launcher_files(
                script_path=script_path,
                launcher_dir=Path(launcher_dir),
                create_desktop_shortcut=create_shortcut,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Launcher Error", str(exc))
            return

        message = f"Created launcher:\n{launcher_path}"
        if shortcut_path:
            message += f"\n\nCreated desktop shortcut:\n{shortcut_path}"
        messagebox.showinfo("Launcher Created", message)
        self.log(message.replace("\n", " "))


def run_cli_launcher_creation() -> int:
    parser = argparse.ArgumentParser(description="Create launcher files for ICO Batch Converter.")
    parser.add_argument("--create-launcher", action="store_true", help="Create an Autorun .bat file and exit.")
    parser.add_argument("--launcher-dir", type=Path, default=None, help="Folder where the Autorun .bat file should be saved.")
    parser.add_argument("--desktop-shortcut", action="store_true", help="Also create a desktop shortcut on Windows.")
    args, _ = parser.parse_known_args()

    if not args.create_launcher:
        return -1

    script_path = Path(__file__).resolve()
    launcher_dir = args.launcher_dir or script_path.parent
    launcher_path, shortcut_path = create_launcher_files(script_path, launcher_dir, args.desktop_shortcut)
    print(f"Created launcher: {launcher_path}")
    if shortcut_path:
        print(f"Created desktop shortcut: {shortcut_path}")
    return 0


def main() -> int:
    launcher_result = run_cli_launcher_creation()
    if launcher_result >= 0:
        return launcher_result

    root = tk.Tk()
    IcoConverterApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
