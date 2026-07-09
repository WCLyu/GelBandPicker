from __future__ import annotations

import importlib
import json
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


APP_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = APP_DIR.parent

REQUIRED_IMPORTS = {
    "numpy": "numpy",
    "pandas": "pandas",
    "PIL": "pillow",
}


def _check_dependencies() -> None:
    missing: list[str] = []
    for import_name, package_name in REQUIRED_IMPORTS.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing.append(package_name)
    if missing:
        requirements = APP_DIR / "requirements.txt"
        raise SystemExit(
            "Missing dependencies: "
            + ", ".join(missing)
            + "\n\nInstall them from the GelBandPicker folder with:\n"
            + f"  {sys.executable} -m pip install -r \"{requirements}\""
        )


_check_dependencies()

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageTk


IMAGE_EXTENSIONS = {".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp"}
DEFAULT_IMAGE = APP_DIR / "example_gel.tif"
DEFAULT_ANNOTATIONS = APP_DIR / "annotation_template.csv"
DEFAULT_OUTDIR = APP_DIR / "outputs"
DEFAULT_LABEL_Y = 80
DEFAULT_TOP_FONT_SIZE = 13
DEFAULT_BOTTOM_FONT_SIZE = 11
DEFAULT_BAND_FONT_SIZE = 11
DEFAULT_WINDOW_WIDTH = 1220
DEFAULT_WINDOW_HEIGHT = 820
DEFAULT_WINDOW_MARGIN = 36
BAND_BUTTONS_PER_ROW = 6
BAND_COLORS = ["#00dc78", "#ffaa00", "#ff465a", "#50a0ff", "#c864ff", "#00d0ff", "#f0d000"]


@dataclass
class LaneSpec:
    lane_index: int
    lane_label: str
    label_top: str
    label_bottom: str
    include: bool = True
    x_center: int | None = None
    note: str = ""


@dataclass
class BandPoint:
    lane_pos: int
    band: str
    y: float
    peak: float = 1.0


def _as_bool(value: object, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not pd.isna(value):
        return bool(value)
    text = str(value).strip().lower()
    if not text or text in {"nan", "none"}:
        return default
    return text not in {"0", "false", "f", "no", "n", "exclude", "ignore"}


def _as_text(value: object, default: str = "") -> str:
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if text.lower() == "nan":
        return default
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text if text else default


def _resolve_column(df: pd.DataFrame, preferred: str, fallbacks: list[str]) -> str | None:
    lower_to_name = {str(col).strip().lower(): col for col in df.columns}
    for name in [preferred] + fallbacks:
        if not name:
            continue
        match = lower_to_name.get(str(name).strip().lower())
        if match is not None:
            return match
    return None


def read_table(path: Path, sheet_name: str) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        try:
            try:
                return pd.read_excel(path, sheet_name=sheet_name or 0)
            except ValueError:
                return pd.read_excel(path, sheet_name=0)
        except ImportError:
            return read_minimal_xlsx(path, sheet_name=sheet_name or 0)
    if suffix == ".xls":
        try:
            try:
                return pd.read_excel(path, sheet_name=sheet_name or 0)
            except ValueError:
                return pd.read_excel(path, sheet_name=0)
        except ImportError as exc:
            raise RuntimeError("Reading .xls needs an Excel reader. Save the table as .xlsx or .csv.") from exc
    if suffix == ".tsv":
        return pd.read_csv(path, sep="\t", dtype=str, keep_default_na=False)
    return pd.read_csv(path, dtype=str, keep_default_na=False)


def _column_index_from_ref(cell_ref: str) -> int:
    letters = ""
    for char in cell_ref:
        if char.isalpha():
            letters += char.upper()
        else:
            break
    if not letters:
        return 0
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def _coerce_xlsx_scalar(value: str) -> object:
    text = value.strip()
    if not text:
        return ""
    try:
        number = float(text)
    except ValueError:
        return text
    return int(number) if number.is_integer() else number


def _xlsx_part_path(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return f"xl/{target}"


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings: list[str] = []
    for item in root.findall("x:si", ns):
        strings.append("".join(node.text or "" for node in item.findall(".//x:t", ns)))
    return strings


def _unique_headers(values: list[object]) -> list[str]:
    headers: list[str] = []
    counts: dict[str, int] = {}
    for idx, value in enumerate(values, start=1):
        base = _as_text(value, f"column_{idx}")
        count = counts.get(base, 0) + 1
        counts[base] = count
        headers.append(base if count == 1 else f"{base}_{count}")
    return headers


def read_minimal_xlsx(path: Path, sheet_name: str | int | None = "Annotations") -> pd.DataFrame:
    ns = {
        "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    with zipfile.ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall("pr:Relationship", ns)}
        sheets = []
        for sheet in workbook.findall("x:sheets/x:sheet", ns):
            sheets.append({"name": sheet.attrib["name"], "rid": sheet.attrib[f"{{{ns['r']}}}id"]})
        if not sheets:
            return pd.DataFrame()
        if sheet_name is None:
            selected = sheets[0]
        elif isinstance(sheet_name, int):
            selected = sheets[sheet_name]
        else:
            selected = next((sheet for sheet in sheets if sheet["name"] == sheet_name), sheets[0])
        sheet_xml = ET.fromstring(zf.read(_xlsx_part_path(rels[selected["rid"]])))
        shared_strings = _read_shared_strings(zf)

    rows: list[list[object]] = []
    for row in sheet_xml.findall("x:sheetData/x:row", ns):
        values_by_col: dict[int, object] = {}
        max_col = -1
        for cell in row.findall("x:c", ns):
            col_idx = _column_index_from_ref(cell.attrib.get("r", ""))
            max_col = max(max_col, col_idx)
            cell_type = cell.attrib.get("t", "")
            if cell_type == "inlineStr":
                value = "".join(node.text or "" for node in cell.findall(".//x:t", ns))
            else:
                value_node = cell.find("x:v", ns)
                raw = value_node.text if value_node is not None and value_node.text is not None else ""
                if cell_type == "s" and raw:
                    value = shared_strings[int(raw)] if int(raw) < len(shared_strings) else ""
                elif cell_type == "b":
                    value = raw == "1"
                else:
                    value = _coerce_xlsx_scalar(raw)
            values_by_col[col_idx] = value
        if max_col >= 0:
            rows.append([values_by_col.get(idx, "") for idx in range(max_col + 1)])
    if not rows:
        return pd.DataFrame()
    headers = _unique_headers(rows[0])
    width = len(headers)
    data = [(row + [""] * width)[:width] for row in rows[1:]]
    return pd.DataFrame(data, columns=headers)


def _xlsx_col_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _sheet_xml(rows: list[list[object]]) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]
    for r_idx, row in enumerate(rows, start=1):
        parts.append(f'<row r="{r_idx}">')
        for c_idx, value in enumerate(row):
            if value is None:
                continue
            try:
                if pd.isna(value):
                    continue
            except (TypeError, ValueError):
                pass
            cell_ref = f"{_xlsx_col_name(c_idx)}{r_idx}"
            if isinstance(value, bool):
                parts.append(f'<c r="{cell_ref}" t="b"><v>{1 if value else 0}</v></c>')
            elif isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
                parts.append(f'<c r="{cell_ref}"><v>{float(value)}</v></c>')
            else:
                parts.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def _df_to_rows(df: pd.DataFrame) -> list[list[object]]:
    rows: list[list[object]] = [list(df.columns)]
    for record in df.to_dict(orient="records"):
        rows.append([record.get(col, "") for col in df.columns])
    return rows


def write_minimal_xlsx(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_names = list(sheets)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
""" + "".join(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for i in range(1, len(sheet_names) + 1)
            ) + "</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        )
        workbook_sheets = "".join(
            f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
            for i, name in enumerate(sheet_names, start=1)
        )
        zf.writestr(
            "xl/workbook.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{workbook_sheets}</sheets></workbook>""",
        )
        workbook_rels = "".join(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, len(sheet_names) + 1)
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{workbook_rels}</Relationships>""",
        )
        for i, name in enumerate(sheet_names, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(_df_to_rows(sheets[name])))


def write_result_xlsx(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    try:
        with pd.ExcelWriter(path) as writer:
            for sheet_name, table in sheets.items():
                table.to_excel(writer, sheet_name=sheet_name, index=False)
    except ImportError:
        write_minimal_xlsx(path, sheets)


def load_annotations(
    path: Path,
    sheet_name: str,
    label_top_col: str,
    label_bottom_col: str,
    lane_label_col: str,
) -> list[LaneSpec]:
    if not path.exists():
        raise FileNotFoundError(f"Annotation table is required and was not found: {path}")

    df = read_table(path, sheet_name)
    if df.empty:
        raise ValueError(f"Annotation table is empty: {path}")

    lane_index_col = _resolve_column(df, "lane_index", ["lane", "lane_no", "index"])
    include_col = _resolve_column(df, "include", ["use", "quantify", "included"])
    lane_col = _resolve_column(df, lane_label_col, ["lane_label", "sample", "sample_id"])
    top_col = _resolve_column(df, label_top_col, ["top_label", "label_top", "sample_name", "sample_id", "sample"])
    bottom_col = _resolve_column(df, label_bottom_col, ["bottom_label", "label_bottom", "condition", "treatment", "group"])
    x_col = _resolve_column(df, "x_center", ["center_x", "x", "lane_center"])
    note_col = _resolve_column(df, "note", ["notes"])

    lanes: list[LaneSpec] = []
    for row_idx, row in df.iterrows():
        lane_index = int(row[lane_index_col]) if lane_index_col and not pd.isna(row[lane_index_col]) else row_idx + 1
        include = _as_bool(row[include_col], True) if include_col else True
        lane_label = _as_text(row[lane_col], str(lane_index)) if lane_col else str(lane_index)
        label_top = _as_text(row[top_col], lane_label) if top_col else lane_label
        label_bottom = _as_text(row[bottom_col], "") if bottom_col else ""
        note = _as_text(row[note_col], "") if note_col else ""
        x_center = None
        if x_col and not pd.isna(row[x_col]):
            try:
                x_center = int(round(float(row[x_col])))
            except (TypeError, ValueError):
                x_center = None
        lanes.append(LaneSpec(lane_index, lane_label, label_top, label_bottom, include, x_center, note))

    lanes.sort(key=lambda lane: lane.lane_index)
    return lanes


def read_gel_image(path: Path, frame: int = 0) -> np.ndarray:
    image = Image.open(path)
    try:
        image.seek(frame)
    except EOFError as exc:
        raise ValueError(f"Image {path} does not contain frame {frame}.") from exc
    arr = np.asarray(image)
    if arr.ndim == 2:
        return arr.astype(float)
    if arr.ndim == 3:
        rgb = np.asarray(image.convert("RGB"), dtype=float)
        return 0.299 * rgb[:, :, 0] + 0.587 * rgb[:, :, 1] + 0.114 * rgb[:, :, 2]
    raise ValueError(f"Unsupported image shape: {arr.shape}")


def display_from_gray(
    gray: np.ndarray,
    invert: bool = False,
    black_percent: float = 0.5,
    white_percent: float = 99.8,
) -> Image.Image:
    finite = gray[np.isfinite(gray)]
    if finite.size:
        black_percent = max(0.0, min(100.0, float(black_percent)))
        white_percent = max(0.0, min(100.0, float(white_percent)))
        if white_percent <= black_percent:
            white_percent = min(100.0, black_percent + 0.1)
        lo, hi = np.percentile(finite, [black_percent, white_percent])
        if hi <= lo:
            lo, hi = float(np.nanmin(finite)), float(np.nanmax(finite))
    else:
        lo, hi = 0.0, 1.0
    scaled = np.clip((gray - lo) / max(hi - lo, 1e-9), 0.0, 1.0)
    out = np.uint8(scaled * 255)
    if invert:
        out = 255 - out
    return Image.fromarray(out, mode="L").convert("RGB")


def analysis_signal(gray: np.ndarray, polarity: str) -> np.ndarray:
    finite = gray[np.isfinite(gray)]
    if finite.size == 0:
        return np.zeros_like(gray, dtype=float)
    if polarity == "dark":
        hi = float(np.percentile(finite, 99.8))
        return np.maximum(hi - gray, 0.0)
    return gray.astype(float)


def smooth_1d(values: np.ndarray, window: int) -> np.ndarray:
    window = max(1, int(window))
    if window <= 1:
        return values.astype(float)
    kernel = np.ones(window, dtype=float) / float(window)
    return np.convolve(values, kernel, mode="same")


def auto_lane_centers(signal: np.ndarray, lane_count: int) -> list[int]:
    height, width = signal.shape
    lane_count = max(1, int(lane_count))
    y1 = int(height * 0.22)
    y2 = int(height * 0.75)
    if y2 <= y1 + 5:
        y1, y2 = 0, height
    roi = signal[y1:y2, :]
    profile = np.nanpercentile(roi, 90, axis=0)
    profile = smooth_1d(profile - np.nanmedian(profile), max(7, width // 180))
    threshold = float(np.nanpercentile(profile, 65))
    min_distance = max(8, int(width / max(lane_count * 2.5, 1)))
    candidates: list[tuple[float, int]] = []
    for x in range(1, width - 1):
        if profile[x] >= threshold and profile[x] >= profile[x - 1] and profile[x] >= profile[x + 1]:
            candidates.append((float(profile[x]), x))
    candidates.sort(reverse=True)
    picked: list[int] = []
    for _score, x in candidates:
        if all(abs(x - current) >= min_distance for current in picked):
            picked.append(x)
        if len(picked) >= lane_count:
            break
    if len(picked) == lane_count:
        return sorted(picked)
    margin = max(12, int(width / max(lane_count + 1, 2) * 0.55))
    return [int(round(x)) for x in np.linspace(margin, width - margin, lane_count)]


def estimate_label_y(signal: np.ndarray) -> int:
    height = signal.shape[0]
    y1 = int(height * 0.04)
    y2 = int(height * 0.35)
    if y2 <= y1 + 5:
        return DEFAULT_LABEL_Y
    roi = signal[y1:y2, :]
    row_score = np.nanpercentile(roi, 96, axis=1) - np.nanpercentile(roi, 5, axis=1)
    row_score = smooth_1d(row_score, max(5, height // 220))
    peak = y1 + int(np.nanargmax(row_score))
    return min(height - 1, peak + max(14, int(height * 0.025)))


def clip_interval(center: int, half_width: int, limit: int) -> tuple[int, int]:
    x1 = max(0, int(round(center - half_width)))
    x2 = min(limit, int(round(center + half_width + 1)))
    return x1, x2


def background_pixels(
    signal: np.ndarray,
    x_center: int,
    y1: int,
    y2: int,
    lane_half_width: int,
    background_gap: int,
    background_width: int,
) -> np.ndarray:
    height, width = signal.shape
    y1 = max(0, min(height, y1))
    y2 = max(y1, min(height, y2))
    left1 = max(0, x_center - lane_half_width - background_gap - background_width)
    left2 = max(0, x_center - lane_half_width - background_gap)
    right1 = min(width, x_center + lane_half_width + background_gap)
    right2 = min(width, x_center + lane_half_width + background_gap + background_width)
    pieces = []
    if left2 > left1:
        pieces.append(signal[y1:y2, left1:left2].ravel())
    if right2 > right1:
        pieces.append(signal[y1:y2, right1:right2].ravel())
    if not pieces:
        return np.asarray([], dtype=float)
    return np.concatenate(pieces)


def font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    size = max(6, int(size))
    names = [
        "arialbd.ttf" if bold else "arial.ttf",
        "Arial Bold.ttf" if bold else "Arial.ttf",
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
        "LiberationSans-Bold.ttf" if bold else "LiberationSans-Regular.ttf",
    ]
    font_dirs = [
        Path("C:/Windows/Fonts"),
        Path("/System/Library/Fonts/Supplemental"),
        Path("/System/Library/Fonts"),
        Path("/Library/Fonts"),
        Path("/usr/share/fonts/truetype/dejavu"),
        Path("/usr/share/fonts/truetype/liberation2"),
        Path("/usr/share/fonts/dejavu"),
        Path(ImageFont.__file__).resolve().parent / "fonts",
        Path.cwd(),
    ]
    candidates = [directory / name for directory in font_dirs for name in names]
    candidates.extend(
        [
            Path("/System/Library/Fonts/Helvetica.ttc"),
            Path("/System/Library/Fonts/Supplemental/Helvetica.ttf"),
            Path("/System/Library/Fonts/Supplemental/Helvetica Bold.ttf"),
        ]
    )
    for candidate in candidates:
        try:
            return ImageFont.truetype(str(candidate), size=size)
        except OSError:
            continue
    try:
        return ImageFont.truetype("DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf", size=size)
    except OSError:
        try:
            return ImageFont.load_default(size=size)
        except TypeError:
            return ImageFont.load_default()


def text_size(draw: ImageDraw.ImageDraw, text: str, text_font: ImageFont.ImageFont) -> tuple[int, int]:
    if not text:
        return 0, 0
    bbox = draw.textbbox((0, 0), text, font=text_font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def draw_centered(
    draw: ImageDraw.ImageDraw,
    x: int,
    y: int,
    text: str,
    text_font: ImageFont.ImageFont,
    fill: tuple[int, int, int],
) -> None:
    if not text:
        return
    width, _height = text_size(draw, text, text_font)
    draw.text((int(x - width / 2), int(y)), text, font=text_font, fill=fill)


class GelGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("GelBandPicker")
        self._center_window(DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT)

        self.image_var = tk.StringVar(value=str(DEFAULT_IMAGE) if DEFAULT_IMAGE.exists() else "")
        self.annotations_var = tk.StringVar(value=str(DEFAULT_ANNOTATIONS) if DEFAULT_ANNOTATIONS.exists() else "")
        self.outdir_var = tk.StringVar(value=str(DEFAULT_OUTDIR))
        self.sheet_var = tk.StringVar(value="Annotations")
        self.top_col_var = tk.StringVar(value="top_label")
        self.bottom_col_var = tk.StringVar(value="bottom_label")
        self.lane_col_var = tk.StringVar(value="lane_label")
        self.polarity_var = tk.StringVar(value="bright")
        self.mode_var = tk.StringVar(value="move")
        self.band_count_var = tk.StringVar(value="3")
        self.current_band_var = tk.StringVar(value="1")
        self.lane_half_width_var = tk.IntVar(value=24)
        self.band_half_height_var = tk.IntVar(value=13)
        self.background_gap_var = tk.IntVar(value=8)
        self.background_width_var = tk.IntVar(value=24)
        self.marker_width_var = tk.IntVar(value=3)
        self.label_y_var = tk.IntVar(value=DEFAULT_LABEL_Y)
        self.top_font_size_var = tk.IntVar(value=DEFAULT_TOP_FONT_SIZE)
        self.bottom_font_size_var = tk.IntVar(value=DEFAULT_BOTTOM_FONT_SIZE)
        self.band_font_size_var = tk.IntVar(value=DEFAULT_BAND_FONT_SIZE)
        self.black_percent_var = tk.DoubleVar(value=0.5)
        self.white_percent_var = tk.DoubleVar(value=99.8)
        self.invert_output_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Move mode: drag the image to pan. Choose another mode to edit.")
        self.spinbox_widgets: dict[str, ttk.Spinbox] = {}

        self.original_gray: np.ndarray | None = None
        self.gray: np.ndarray | None = None
        self.signal: np.ndarray | None = None
        self.display: Image.Image | None = None
        self.tk_image: ImageTk.PhotoImage | None = None
        self.crop_box_original: tuple[int, int, int, int] | None = None
        self.lanes: list[LaneSpec] = []
        self.centers: list[int] = []
        self.points: dict[tuple[int, str], BandPoint] = {}
        self.legend_y_overrides: dict[str, float] = {}
        self.scale = 0.5
        self.crop_start: tuple[int, int] | None = None
        self.crop_current: tuple[int, int] | None = None
        self.drag_lane: int | None = None
        self.drag_legend_band: str | None = None
        self.panning = False
        self.undo_stack: list[dict[str, object]] = []
        self.redo_stack: list[dict[str, object]] = []

        self._build_ui()
        self._refresh_band_buttons()

    def _center_window(self, width: int, height: int) -> None:
        self.root.update_idletasks()
        area_x, area_y, area_width, area_height = self._available_work_area()
        max_width = max(760, area_width - DEFAULT_WINDOW_MARGIN * 2)
        max_height = max(620, area_height - DEFAULT_WINDOW_MARGIN * 2)
        width = min(width, max_width)
        height = min(height, max_height)
        x = area_x + max(0, int((area_width - width) / 2))
        y = area_y + max(0, int((area_height - height) / 2))
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _available_work_area(self) -> tuple[int, int, int, int]:
        if sys.platform.startswith("win"):
            try:
                import ctypes
                from ctypes import wintypes

                rect = wintypes.RECT()
                ok = ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
                if ok:
                    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
            except Exception:
                pass
        return 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()

    def _build_ui(self) -> None:
        ttk_style = ttk.Style(self.root)
        ttk_style.configure("Primary.TButton", font=("TkDefaultFont", 11, "bold"), padding=(24, 12))

        top = ttk.Frame(self.root, padding=8)
        top.pack(side=tk.TOP, fill=tk.X)
        top.columnconfigure(0, weight=1)

        path_box = ttk.LabelFrame(top, text="Files")
        path_box.grid(row=0, column=0, columnspan=4, sticky="ew")
        path_box.columnconfigure(1, weight=1)
        self._path_row(path_box, 0, "Image", self.image_var, self._browse_image)
        self._path_row(path_box, 1, "Annotations", self.annotations_var, self._browse_annotations)
        self._path_row(path_box, 2, "Output", self.outdir_var, self._browse_outdir)
        ttk.Button(
            path_box,
            text="LOAD IMAGE",
            command=self.load_image,
            padding=(18, 18),
        ).grid(row=0, column=3, rowspan=3, sticky="nsew", padx=(12, 8), pady=4)

        workflow = ttk.LabelFrame(top, text="Workflow")
        workflow.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        workflow.columnconfigure(11, weight=1)
        row0 = ttk.Frame(workflow)
        row0.grid(row=0, column=0, sticky="ew", padx=6, pady=(2, 1))
        row1 = ttk.Frame(workflow)
        row1.grid(row=1, column=0, sticky="ew", padx=6, pady=(1, 3))
        row1.columnconfigure(7, weight=1)
        ttk.Label(row0, text="Mode").pack(side=tk.LEFT, padx=(0, 8))
        for text, value in [
            ("Move", "move"),
            ("Pick bands", "band"),
            ("Crop", "crop"),
            ("Edit layout", "lane"),
        ]:
            ttk.Radiobutton(
                row0,
                text=text,
                value=value,
                variable=self.mode_var,
                command=self.on_mode_changed,
            ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(row0, text="Apply crop", command=self.apply_crop).pack(side=tk.LEFT, padx=(4, 4))
        ttk.Button(row0, text="Auto lanes", command=self.auto_detect_lanes).pack(side=tk.LEFT, padx=4)
        ttk.Button(row0, text="Reset", command=self.reset_crop).pack(side=tk.LEFT, padx=4)
        ttk.Button(row1, text="Undo", command=self.undo).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(row1, text="Redo", command=self.redo).pack(side=tk.LEFT, padx=4)
        ttk.Button(row1, text="Zoom -", command=lambda: self.set_scale(self.scale / 1.25)).pack(side=tk.LEFT, padx=(12, 4))
        ttk.Button(row1, text="Zoom +", command=lambda: self.set_scale(self.scale * 1.25)).pack(side=tk.LEFT, padx=4)
        ttk.Label(row1, textvariable=self.status_var).pack(side=tk.LEFT, padx=(14, 8), fill=tk.X, expand=True)

        ann = ttk.LabelFrame(top, text="Annotation")
        ann.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        for label, var, width in [
            ("Sheet", self.sheet_var, 12),
            ("Top label", self.top_col_var, 12),
            ("Bottom label", self.bottom_col_var, 12),
            ("Lane label", self.lane_col_var, 12),
        ]:
            ttk.Label(ann, text=label).pack(side=tk.LEFT, padx=(8 if label == "Sheet" else 0, 0))
            ttk.Entry(ann, textvariable=var, width=width).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(ann, text="Polarity").pack(side=tk.LEFT)
        ttk.Combobox(
            ann,
            textvariable=self.polarity_var,
            values=("bright", "dark"),
            state="readonly",
            width=7,
        ).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(ann, text="Save session", command=self.save_session).pack(side=tk.LEFT, padx=(14, 5))
        ttk.Button(ann, text="Load session", command=self.load_session).pack(side=tk.LEFT, padx=5)

        band_box = ttk.LabelFrame(top, text="Band picking and ROI")
        band_box.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        band_settings = ttk.Frame(band_box)
        band_settings.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(3, 0))
        self._band_count_entry(band_settings)
        self.band_frame = ttk.Frame(band_settings)
        self.band_frame.pack(side=tk.LEFT, padx=(0, 14))
        self._spin(band_settings, "Band half width", self.lane_half_width_var, 3, 120, 5, self._marker_changed)
        self._spin(band_settings, "Band half height", self.band_half_height_var, 2, 80, 5, self._marker_changed)
        self._spin(band_settings, "Marker line width", self.marker_width_var, 1, 16, 4, self._marker_changed)

        style = ttk.LabelFrame(top, text="Style and display")
        style.grid(row=4, column=0, columnspan=3, sticky="ew", padx=(0, 8), pady=(8, 0))
        style_fonts = ttk.Frame(style)
        style_fonts.pack(side=tk.TOP, fill=tk.X, padx=6, pady=(2, 1))
        style_display = ttk.Frame(style)
        style_display.pack(side=tk.TOP, fill=tk.X, padx=6, pady=(1, 3))
        self._spin(style_fonts, "Label y", self.label_y_var, 0, 5000, 6, self.redraw)
        self._spin(style_fonts, "Top font", self.top_font_size_var, 6, 40, 4, self.redraw)
        self._spin(style_fonts, "Bottom font", self.bottom_font_size_var, 6, 40, 4, self.redraw)
        self._spin(style_fonts, "Band font", self.band_font_size_var, 6, 40, 4, self.redraw)
        ttk.Label(style_display, text="Black").pack(side=tk.LEFT)
        tk.Scale(
            style_display,
            from_=0.0,
            to=20.0,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            length=130,
            variable=self.black_percent_var,
            command=self._display_changed,
        ).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(style_display, text="White").pack(side=tk.LEFT)
        tk.Scale(
            style_display,
            from_=80.0,
            to=100.0,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            length=130,
            variable=self.white_percent_var,
            command=self._display_changed,
        ).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Checkbutton(
            style_display,
            text="Invert",
            variable=self.invert_output_var,
            command=self._display_changed,
        ).pack(side=tk.LEFT, padx=(0, 8))

        action_row = ttk.Frame(top)
        action_row.grid(row=4, column=3, sticky="nse", pady=(8, 0))
        ttk.Button(
            action_row,
            text="RUN QUANTIFICATION",
            command=self.run_quantification,
            style="Primary.TButton",
            width=24,
        ).pack(side=tk.RIGHT, fill=tk.Y)

        main = ttk.Frame(self.root)
        main.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(main, background="#111111")
        self.hbar = ttk.Scrollbar(main, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.vbar = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        self.canvas.bind("<ButtonPress-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)
        self.canvas.bind("<Double-Button-1>", self.on_double_click)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self.on_shift_mouse_wheel)
        self.canvas.bind("<Button-4>", self.on_mouse_wheel)
        self.canvas.bind("<Button-5>", self.on_mouse_wheel)
        self.canvas.bind("<Shift-Button-4>", self.on_shift_mouse_wheel)
        self.canvas.bind("<Shift-Button-5>", self.on_shift_mouse_wheel)
        self.root.bind_all("<Control-z>", lambda _event: self.undo())
        self.root.bind_all("<Control-y>", lambda _event: self.redo())
        self.root.bind_all("<Control-Shift-Z>", lambda _event: self.redo())
        self.root.bind_all("<Command-z>", lambda _event: self.undo())
        self.root.bind_all("<Command-Shift-Z>", lambda _event: self.redo())

    def _path_row(self, parent: ttk.Frame, row: int, label: str, var: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6, pady=2)
        ttk.Button(parent, text="Browse", command=command).grid(row=row, column=2, sticky="w", pady=2)
        parent.columnconfigure(1, weight=1)

    def _spin(
        self,
        parent: ttk.Frame,
        label: str,
        var: tk.IntVar,
        from_: int,
        to: int,
        width: int,
        command=None,
    ) -> None:
        ttk.Label(parent, text=label).pack(side=tk.LEFT)
        spinbox = ttk.Spinbox(parent, from_=from_, to=to, textvariable=var, width=width, command=command)
        self.spinbox_widgets[str(var)] = spinbox
        spinbox.pack(side=tk.LEFT, padx=(4, 12))
        if command is not None:
            spinbox.bind("<Return>", lambda _event: command())
            spinbox.bind("<FocusOut>", lambda _event: command())

    def _band_count_entry(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Bands").pack(side=tk.LEFT)
        entry = ttk.Entry(parent, textvariable=self.band_count_var, width=4)
        entry.pack(side=tk.LEFT, padx=(4, 12))
        entry.bind("<Return>", lambda _event: self._on_band_count_changed())
        entry.bind("<FocusOut>", lambda _event: self._on_band_count_changed())

    def _band_count(self) -> int:
        try:
            return max(1, int(str(self.band_count_var.get()).strip()))
        except (TypeError, ValueError, tk.TclError):
            return 1

    def _int_value(
        self,
        var: tk.IntVar,
        default: int,
        minimum: int | None = None,
        maximum: int | None = None,
    ) -> int:
        widget = self.spinbox_widgets.get(str(var))
        raw_value = widget.get() if widget is not None else var.get()
        try:
            value = int(float(str(raw_value).strip()))
        except (TypeError, ValueError, tk.TclError):
            value = default
        if minimum is not None:
            value = max(minimum, value)
        if maximum is not None:
            value = min(maximum, value)
        try:
            var.set(value)
            if widget is not None and widget.get() != str(value):
                widget.delete(0, tk.END)
                widget.insert(0, str(value))
        except tk.TclError:
            pass
        return value

    def _sync_numeric_settings(self) -> None:
        self._int_value(self.lane_half_width_var, 24, 3, 120)
        self._int_value(self.band_half_height_var, 13, 2, 80)
        self._int_value(self.marker_width_var, 3, 1, 16)
        self._int_value(self.label_y_var, DEFAULT_LABEL_Y, 0, 5000)
        self._int_value(self.top_font_size_var, DEFAULT_TOP_FONT_SIZE, 6, 40)
        self._int_value(self.bottom_font_size_var, DEFAULT_BOTTOM_FONT_SIZE, 6, 40)
        self._int_value(self.band_font_size_var, DEFAULT_BAND_FONT_SIZE, 6, 40)

    def _browse_image(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=str(WORKSPACE_DIR),
            filetypes=[("Images", "*.tif *.tiff *.png *.jpg *.jpeg *.bmp"), ("All files", "*.*")],
        )
        if path:
            self.image_var.set(path)

    def _browse_annotations(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=str(WORKSPACE_DIR),
            filetypes=[("Tables", "*.xlsx *.xls *.csv *.tsv"), ("All files", "*.*")],
        )
        if path:
            self.annotations_var.set(path)

    def _browse_outdir(self) -> None:
        path = filedialog.askdirectory(initialdir=str(APP_DIR))
        if path:
            self.outdir_var.set(path)

    def _on_band_count_changed(self) -> None:
        self.band_count_var.set(str(self._band_count()))
        self._refresh_band_buttons()
        self.redraw()

    def _refresh_band_buttons(self) -> None:
        for child in self.band_frame.winfo_children():
            child.destroy()
        count = self._band_count()
        for idx in range(1, count + 1):
            row = (idx - 1) // BAND_BUTTONS_PER_ROW
            col = (idx - 1) % BAND_BUTTONS_PER_ROW
            ttk.Radiobutton(
                self.band_frame,
                text=f"Band {idx}",
                value=str(idx),
                variable=self.current_band_var,
            ).grid(row=row, column=col, sticky="w", padx=3, pady=1)
        try:
            if int(self.current_band_var.get()) > count:
                self.current_band_var.set(str(count))
        except ValueError:
            self.current_band_var.set("1")

    def _marker_changed(self) -> None:
        self.redraw()

    def _display_changed(self, *_args) -> None:
        if self.gray is None:
            return
        self._refresh_display(reset_scale=False)
        self.redraw()

    def _refresh_display(self, reset_scale: bool) -> None:
        if self.gray is None:
            return
        self.display = display_from_gray(
            self.gray,
            invert=bool(self.invert_output_var.get()),
            black_percent=float(self.black_percent_var.get()),
            white_percent=float(self.white_percent_var.get()),
        )
        if reset_scale:
            width, height = self.display.size
            self.scale = min(1.0, 1200 / max(width, 1), 700 / max(height, 1))

    def on_mode_changed(self) -> None:
        self.crop_start = None
        self.crop_current = None
        self.drag_lane = None
        self.drag_legend_band = None
        self.panning = False
        mode = self.mode_var.get()
        if mode == "move":
            self.status_var.set("Move mode: drag the image to pan. No labels or bands will be added.")
        elif mode == "crop":
            self.status_var.set("Crop mode: drag a rectangle, then click Apply crop.")
        elif mode == "lane":
            self.status_var.set("Edit layout: drag lane lines or band legends. Right-click a legend to reset it.")
        else:
            self.status_var.set("Pick bands: choose Band 1/2/3, then click the matching bands.")
        self.redraw()

    def load_image(self) -> None:
        try:
            image_path = Path(self.image_var.get()).expanduser()
            if not image_path.exists():
                raise FileNotFoundError(f"Image not found: {image_path}")
            self.original_gray = read_gel_image(image_path)
            self.crop_box_original = None
            self._set_working_gray(self.original_gray)
            self._load_lanes()
            self.auto_detect_lanes(record=False)
            self.label_y_var.set(int(estimate_label_y(self.signal)) if self.signal is not None else DEFAULT_LABEL_Y)
            self.points.clear()
            self.legend_y_overrides.clear()
            self._clear_history()
            self.status_var.set(f"Loaded {image_path.name}. Use Crop, Edit layout, or Pick bands.")
            self.redraw()
        except Exception as exc:
            messagebox.showerror("Load failed", str(exc))

    def _load_lanes(self) -> None:
        ann_text = self.annotations_var.get().strip()
        if not ann_text:
            raise ValueError("Annotation table is required. Choose an annotation .csv/.tsv/.xlsx file.")
        ann_path = Path(ann_text).expanduser()
        self.lanes = load_annotations(
            ann_path,
            self.sheet_var.get().strip() or "Annotations",
            self.top_col_var.get().strip() or "top_label",
            self.bottom_col_var.get().strip() or "bottom_label",
            self.lane_col_var.get().strip() or "lane_label",
        )

    def _set_working_gray(self, gray: np.ndarray) -> None:
        self.gray = gray.astype(float)
        self.signal = analysis_signal(self.gray, self.polarity_var.get())
        self._refresh_display(reset_scale=True)

    def auto_detect_lanes(self, record: bool = True) -> None:
        if self.signal is None:
            return
        if record:
            self._push_undo()
        explicit = [lane.x_center for lane in self.lanes]
        if explicit and all(center is not None for center in explicit):
            self.centers = [int(center) for center in explicit if center is not None]
        else:
            self.centers = auto_lane_centers(self.signal, len(self.lanes))
        self._sync_lane_count_to_centers()
        self.status_var.set(f"Lane centers ready: {len(self.centers)} lanes.")
        self.redraw()

    def _sync_lane_count_to_centers(self) -> None:
        while len(self.lanes) < len(self.centers):
            idx = len(self.lanes) + 1
            self.lanes.append(LaneSpec(idx, str(idx), str(idx), "", True, None, "added in GUI"))
        if len(self.lanes) > len(self.centers):
            self.lanes = self.lanes[: len(self.centers)]

    def _manual_state(self) -> dict[str, object]:
        return {
            "lanes": [asdict(lane) for lane in self.lanes],
            "centers": list(self.centers),
            "points": [asdict(point) for point in self.points.values()],
            "legend_y_overrides": dict(self.legend_y_overrides),
        }

    def _restore_manual_state(self, state: dict[str, object]) -> None:
        self.lanes = [LaneSpec(**lane) for lane in state.get("lanes", [])]
        self.centers = [int(center) for center in state.get("centers", [])]
        self.legend_y_overrides = {
            str(band): float(y)
            for band, y in dict(state.get("legend_y_overrides", {})).items()
        }
        self.points = {}
        for item in state.get("points", []):
            point = BandPoint(int(item["lane_pos"]), str(item["band"]), float(item["y"]), float(item.get("peak", 1.0)))
            self.points[(point.lane_pos, point.band)] = point
        self.redraw()

    def _push_undo(self) -> None:
        self.undo_stack.append(self._manual_state())
        if len(self.undo_stack) > 100:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _clear_history(self) -> None:
        self.undo_stack.clear()
        self.redo_stack.clear()

    def undo(self) -> None:
        if not self.undo_stack:
            self.status_var.set("Nothing to undo.")
            return
        self.redo_stack.append(self._manual_state())
        self._restore_manual_state(self.undo_stack.pop())
        self.status_var.set("Undid last manual edit.")

    def redo(self) -> None:
        if not self.redo_stack:
            self.status_var.set("Nothing to redo.")
            return
        self.undo_stack.append(self._manual_state())
        self._restore_manual_state(self.redo_stack.pop())
        self.status_var.set("Redid manual edit.")

    def set_scale(self, scale: float) -> None:
        self.scale = max(0.05, min(4.0, scale))
        self.redraw()

    def _wheel_units(self, event: tk.Event) -> int:
        delta = getattr(event, "delta", 0)
        if delta:
            return -1 if delta > 0 else 1
        button = getattr(event, "num", 0)
        if button == 4:
            return -1
        if button == 5:
            return 1
        return 0

    def on_mouse_wheel(self, event: tk.Event) -> None:
        units = self._wheel_units(event)
        if units:
            self.canvas.yview_scroll(units * 4, "units")

    def on_shift_mouse_wheel(self, event: tk.Event) -> None:
        units = self._wheel_units(event)
        if units:
            self.canvas.xview_scroll(units * 4, "units")

    def image_xy_from_event(self, event: tk.Event) -> tuple[int, int]:
        x = self.canvas.canvasx(event.x) / self.scale
        y = self.canvas.canvasy(event.y) / self.scale
        return int(round(x)), int(round(y))

    def nearest_lane(self, x: int) -> tuple[int | None, int]:
        if not self.centers:
            return None, 10**9
        idx = min(range(len(self.centers)), key=lambda lane_pos: abs(self.centers[lane_pos] - x))
        return idx, abs(self.centers[idx] - x)

    def on_left_down(self, event: tk.Event) -> None:
        if self.display is None:
            return
        x, y = self.image_xy_from_event(event)
        mode = self.mode_var.get()
        if mode == "move":
            self.canvas.scan_mark(event.x, event.y)
            self.panning = True
            return
        if mode == "crop":
            self.crop_start = (x, y)
            self.crop_current = (x, y)
            self.redraw()
            return
        if mode == "lane":
            legend_band = self.nearest_legend(x, y)
            if legend_band is not None:
                self._push_undo()
                self.drag_legend_band = legend_band
                self.legend_y_overrides[legend_band] = float(y)
                self.status_var.set(f"Drag band {legend_band} legend. Right-click it to return to automatic average.")
                self.redraw()
                return
            lane_pos, distance = self.nearest_lane(x)
            if lane_pos is not None and distance <= max(12, int(10 / max(self.scale, 0.1))):
                self._push_undo()
                self.drag_lane = lane_pos
                self.status_var.set(f"Drag lane {self.lanes[lane_pos].lane_label}. Double-click empty space to add a lane.")
            return
        if mode == "band":
            self.set_band_point(x, y)

    def on_left_drag(self, event: tk.Event) -> None:
        if self.display is None:
            return
        if self.mode_var.get() == "move" and self.panning:
            self.canvas.scan_dragto(event.x, event.y, gain=1)
            return
        x, y = self.image_xy_from_event(event)
        if self.mode_var.get() == "crop" and self.crop_start is not None:
            self.crop_current = (x, y)
            self.redraw()
            return
        if self.mode_var.get() == "lane" and self.drag_legend_band is not None:
            height = self.signal.shape[0] if self.signal is not None else max(1, int(y))
            self.legend_y_overrides[self.drag_legend_band] = float(max(0, min(height - 1, y)))
            self.redraw()
            return
        if self.mode_var.get() == "lane" and self.drag_lane is not None:
            width = self.signal.shape[1] if self.signal is not None else 1
            min_x = self.centers[self.drag_lane - 1] + 4 if self.drag_lane > 0 else 0
            max_x = self.centers[self.drag_lane + 1] - 4 if self.drag_lane < len(self.centers) - 1 else width - 1
            self.centers[self.drag_lane] = int(max(min_x, min(max_x, x)))
            self.redraw()

    def on_left_up(self, _event: tk.Event) -> None:
        if self.drag_legend_band is not None:
            self.status_var.set(f"Band {self.drag_legend_band} legend adjusted.")
        if self.drag_lane is not None:
            self.status_var.set("Lane center adjusted.")
        self.drag_legend_band = None
        self.drag_lane = None
        self.panning = False

    def on_double_click(self, event: tk.Event) -> None:
        if self.mode_var.get() != "lane" or self.signal is None:
            return
        x, y = self.image_xy_from_event(event)
        legend_band = self.nearest_legend(x, y)
        if legend_band is not None:
            if legend_band in self.legend_y_overrides:
                self._push_undo()
                self.legend_y_overrides.pop(legend_band, None)
                self.status_var.set(f"Band {legend_band} legend returned to automatic average.")
                self.redraw()
            return
        x = int(max(0, min(self.signal.shape[1] - 1, x)))
        insert_at = 0
        while insert_at < len(self.centers) and self.centers[insert_at] < x:
            insert_at += 1
        self._push_undo()
        self.centers.insert(insert_at, x)
        lane_idx = insert_at + 1
        self.lanes.insert(insert_at, LaneSpec(lane_idx, str(lane_idx), str(lane_idx), "", True, None, "added in GUI"))
        self._renumber_added_lanes()
        self._shift_points_after_insert(insert_at)
        self.status_var.set(f"Added lane at x={x}.")
        self.redraw()

    def on_right_click(self, event: tk.Event) -> None:
        if self.display is None:
            return
        x, _y = self.image_xy_from_event(event)
        if self.mode_var.get() == "move":
            return
        if self.mode_var.get() == "crop":
            self.crop_start = None
            self.crop_current = None
            self.redraw()
            return
        if self.mode_var.get() == "lane":
            legend_band = self.nearest_legend(x, _y)
            if legend_band is not None:
                if legend_band in self.legend_y_overrides:
                    self._push_undo()
                    self.legend_y_overrides.pop(legend_band, None)
                    self.status_var.set(f"Band {legend_band} legend returned to automatic average.")
                    self.redraw()
                return
        lane_pos, distance = self.nearest_lane(x)
        if lane_pos is None:
            return
        if self.mode_var.get() == "lane" and distance <= max(14, int(12 / max(self.scale, 0.1))):
            self._push_undo()
            del self.centers[lane_pos]
            del self.lanes[lane_pos]
            self._remove_lane_points(lane_pos)
            self._renumber_added_lanes()
            self.status_var.set("Removed lane.")
            self.redraw()
            return
        if self.mode_var.get() != "band":
            return
        band = str(self.current_band_var.get())
        key = (lane_pos, band)
        if key in self.points:
            self._push_undo()
            del self.points[key]
            self.status_var.set(f"Deleted band {band}, lane {self.lanes[lane_pos].lane_label}.")
            self.redraw()

    def _renumber_added_lanes(self) -> None:
        for idx, lane in enumerate(self.lanes, start=1):
            if lane.note == "added in GUI":
                lane.lane_index = idx
                lane.lane_label = str(idx)
                lane.label_top = str(idx)

    def _shift_points_after_insert(self, insert_at: int) -> None:
        shifted: dict[tuple[int, str], BandPoint] = {}
        for (lane_pos, band), point in self.points.items():
            new_lane_pos = lane_pos + 1 if lane_pos >= insert_at else lane_pos
            point.lane_pos = new_lane_pos
            shifted[(new_lane_pos, band)] = point
        self.points = shifted

    def _remove_lane_points(self, removed: int) -> None:
        shifted: dict[tuple[int, str], BandPoint] = {}
        for (lane_pos, band), point in self.points.items():
            if lane_pos == removed:
                continue
            new_lane_pos = lane_pos - 1 if lane_pos > removed else lane_pos
            point.lane_pos = new_lane_pos
            shifted[(new_lane_pos, band)] = point
        self.points = shifted

    def set_band_point(self, x: int, y: int) -> None:
        lane_pos, _distance = self.nearest_lane(x)
        if lane_pos is None:
            return
        band = str(self.current_band_var.get())
        refined_y, peak = self.refine_peak(lane_pos, y)
        self._push_undo()
        self.points[(lane_pos, band)] = BandPoint(lane_pos, band, float(refined_y), float(peak))
        lane = self.lanes[lane_pos]
        self.status_var.set(f"Set band {band}, lane {lane.lane_label}: y={refined_y}.")
        self.redraw()

    def refine_peak(self, lane_pos: int, y: int) -> tuple[int, float]:
        if self.signal is None:
            return y, 0.0
        height, width = self.signal.shape
        x_center = self.centers[lane_pos]
        band_half = max(2, int(self.band_half_height_var.get()))
        y1 = max(0, y - max(4, band_half))
        y2 = min(height, y + max(4, band_half) + 1)
        x1, x2 = clip_interval(x_center, int(self.lane_half_width_var.get()), width)
        lane_profile = np.nanmean(self.signal[y1:y2, x1:x2], axis=1)
        bg = background_pixels(
            self.signal,
            x_center,
            y1,
            y2,
            int(self.lane_half_width_var.get()),
            int(self.background_gap_var.get()),
            int(self.background_width_var.get()),
        )
        bg_median = float(np.nanmedian(bg)) if bg.size else 0.0
        profile = lane_profile - bg_median
        if profile.size == 0 or not np.isfinite(profile).any():
            return y, 0.0
        offset = int(np.nanargmax(profile))
        return y1 + offset, float(profile[offset])

    def apply_crop(self) -> None:
        if self.gray is None or self.crop_start is None or self.crop_current is None:
            messagebox.showinfo("Crop", "Switch to Crop mode and drag a rectangle first.")
            return
        x0, y0 = self.crop_start
        x1, y1 = self.crop_current
        left, right = sorted([max(0, x0), max(0, x1)])
        top, bottom = sorted([max(0, y0), max(0, y1)])
        height, width = self.gray.shape
        left, right = max(0, left), min(width, right)
        top, bottom = max(0, top), min(height, bottom)
        if right - left < 20 or bottom - top < 20:
            messagebox.showwarning("Crop", "Crop area is too small.")
            return

        origin_x = self.crop_box_original[0] if self.crop_box_original else 0
        origin_y = self.crop_box_original[1] if self.crop_box_original else 0
        self.crop_box_original = (origin_x + left, origin_y + top, origin_x + right, origin_y + bottom)
        self._set_working_gray(self.gray[top:bottom, left:right])
        keep = [idx for idx, center in enumerate(self.centers) if left <= center < right]
        old_to_new = {old: new for new, old in enumerate(keep)}
        self.centers = [self.centers[idx] - left for idx in keep]
        self.lanes = [self.lanes[idx] for idx in keep]
        new_points: dict[tuple[int, str], BandPoint] = {}
        for (lane_pos, band), point in self.points.items():
            if lane_pos not in old_to_new:
                continue
            if top <= point.y < bottom:
                new_lane_pos = old_to_new[lane_pos]
                new_point = BandPoint(new_lane_pos, band, point.y - top, point.peak)
                new_points[(new_lane_pos, band)] = new_point
        self.points = new_points
        self.legend_y_overrides = {
            band: y - top
            for band, y in self.legend_y_overrides.items()
            if top <= y < bottom
        }
        self.label_y_var.set(max(0, int(self.label_y_var.get()) - top))
        self._clear_history()
        self.crop_start = None
        self.crop_current = None
        self.status_var.set("Crop applied. Adjust lanes if needed.")
        self.redraw()

    def reset_crop(self) -> None:
        if self.original_gray is None:
            return
        self.crop_box_original = None
        self._set_working_gray(self.original_gray)
        self._load_lanes()
        self.auto_detect_lanes(record=False)
        self.label_y_var.set(int(estimate_label_y(self.signal)) if self.signal is not None else DEFAULT_LABEL_Y)
        self.points.clear()
        self.legend_y_overrides.clear()
        self._clear_history()
        self.crop_start = None
        self.crop_current = None
        self.status_var.set("Reset to the original image and cleared manual picks.")
        self.redraw()

    def redraw(self) -> None:
        self.canvas.delete("all")
        if self.display is None:
            return
        self._sync_numeric_settings()
        width, height = self.display.size
        scaled = self.display.resize((max(1, int(width * self.scale)), max(1, int(height * self.scale))))
        self.tk_image = ImageTk.PhotoImage(scaled)
        self.canvas.create_image(0, 0, image=self.tk_image, anchor=tk.NW)
        self.canvas.configure(scrollregion=(0, 0, scaled.width, scaled.height))

        label_y = self._int_value(self.label_y_var, DEFAULT_LABEL_Y, 0, 5000)
        top_font_size = self._int_value(self.top_font_size_var, DEFAULT_TOP_FONT_SIZE, 6, 40)
        bottom_font_size = self._int_value(self.bottom_font_size_var, DEFAULT_BOTTOM_FONT_SIZE, 6, 40)
        band_font_size = self._int_value(self.band_font_size_var, DEFAULT_BAND_FONT_SIZE, 6, 40)
        top_text_y = label_y * self.scale
        bottom_text_y = (label_y + top_font_size + 4) * self.scale
        lane_tick_y1 = (label_y + top_font_size + bottom_font_size + 12) * self.scale
        lane_tick_y2 = (label_y + top_font_size + bottom_font_size + 28) * self.scale
        top_canvas_font = ("Arial", max(6, int(round(top_font_size * self.scale))), "bold")
        bottom_canvas_font = ("Arial", max(6, int(round(bottom_font_size * self.scale))))
        band_canvas_font = ("Arial", max(6, int(round(band_font_size * self.scale))), "bold")

        for idx, x_center in enumerate(self.centers):
            lane = self.lanes[idx]
            color = "#00e5ff" if lane.include else "#aaaaaa"
            sx = x_center * self.scale
            if self.mode_var.get() == "lane":
                self.canvas.create_line(sx, 0, sx, height * self.scale, fill=color, width=1, dash=(4, 4))
            self.canvas.create_text(sx, top_text_y, text=lane.label_top, fill=color, font=top_canvas_font)
            if lane.label_bottom:
                self.canvas.create_text(sx, bottom_text_y, text=lane.label_bottom, fill=color, font=bottom_canvas_font)
            self.canvas.create_line(sx, lane_tick_y1, sx, lane_tick_y2, fill=color, width=2)

        for (lane_pos, band), point in self.points.items():
            if lane_pos >= len(self.centers):
                continue
            color = self.color_for_band(band)
            x_center = self.centers[lane_pos]
            y = int(round(point.y))
            marker_half = max(2, int(self.lane_half_width_var.get()))
            marker_width = max(1, int(round(int(self.marker_width_var.get()) * self.scale)))
            half_h = int(self.band_half_height_var.get())
            x1 = (x_center - marker_half) * self.scale
            x2 = (x_center + marker_half) * self.scale
            y1 = (y - half_h) * self.scale
            y2 = (y + half_h) * self.scale
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=marker_width)
            self.canvas.create_line(x1, y * self.scale, x2, y * self.scale, fill=color, width=marker_width)
            self.canvas.create_text(x2 + 7, y * self.scale, text=band, fill=color, anchor=tk.W, font=band_canvas_font)

        self._draw_canvas_band_legend(band_canvas_font)

        if self.crop_start is not None and self.crop_current is not None:
            x0, y0 = self.crop_start
            x1, y1 = self.crop_current
            self.canvas.create_rectangle(
                x0 * self.scale,
                y0 * self.scale,
                x1 * self.scale,
                y1 * self.scale,
                outline="#ffea00",
                width=2,
                dash=(6, 4),
            )

    def _band_sort_key(self, band: str) -> tuple[int, object]:
        try:
            return 0, int(band)
        except ValueError:
            return 1, band

    def _band_legend_entries(self) -> list[tuple[str, float]]:
        by_band: dict[str, list[float]] = {}
        for (lane_pos, band), point in self.points.items():
            if lane_pos >= len(self.centers):
                continue
            band_name = str(point.band or band)
            by_band.setdefault(band_name, []).append(float(point.y))
        entries: list[tuple[str, float]] = []
        for band, ys in by_band.items():
            if not ys:
                continue
            entries.append((band, self.legend_y_overrides.get(band, float(np.mean(ys)))))
        return sorted(entries, key=lambda item: self._band_sort_key(item[0]))

    def _band_legend_left_x(self, entries: list[tuple[str, float]], font_size: int, scaled: bool) -> int:
        if not entries or not self.centers:
            return 4
        first_lane_x = min(self.centers)
        lane_half = self._int_value(self.lane_half_width_var, 24, 3, 120)
        max_chars = max(len(f"band {band}") for band, _y in entries)
        label_width = int(max_chars * font_size * 0.62)
        if scaled:
            return max(4, int(round((first_lane_x - lane_half) * self.scale)) - label_width - 42)
        return max(4, first_lane_x - lane_half - label_width - 42)

    def nearest_legend(self, x: int, y: int) -> str | None:
        entries = self._band_legend_entries()
        if not entries or not self.centers:
            return None
        font_size = self._int_value(self.band_font_size_var, DEFAULT_BAND_FONT_SIZE, 6, 40)
        left_x = self._band_legend_left_x(entries, font_size, scaled=False)
        hits: list[tuple[float, str]] = []
        for band, legend_y in entries:
            label = f"band {band}"
            width = int(len(label) * font_size * 0.62) + 42
            if left_x - 8 <= x <= left_x + width + 8 and abs(y - legend_y) <= max(10, font_size):
                hits.append((abs(y - legend_y), band))
        return min(hits)[1] if hits else None

    def _draw_canvas_band_legend(self, band_canvas_font: tuple[str, int, str]) -> None:
        entries = self._band_legend_entries()
        if not entries or not self.centers:
            return
        font_size = max(6, int(round(self._int_value(self.band_font_size_var, DEFAULT_BAND_FONT_SIZE, 6, 40) * self.scale)))
        left_x = self._band_legend_left_x(entries, font_size, scaled=True)
        for band, y in entries:
            color = self.color_for_band(band)
            sy = y * self.scale
            label = f"band {band}"
            approx_width = int(len(label) * font_size * 0.62)
            self.canvas.create_text(left_x, sy, text=label, fill=color, anchor=tk.W, font=band_canvas_font)
            line_x1 = left_x + approx_width + 8
            line_x2 = line_x1 + 24
            self.canvas.create_line(line_x1, sy, line_x2, sy, fill=color, width=2)

    def color_for_band(self, band: str) -> str:
        try:
            idx = int(band) - 1
        except ValueError:
            idx = 0
        return BAND_COLORS[idx % len(BAND_COLORS)]

    def run_quantification(self) -> None:
        try:
            self.root.focus_set()
            self.root.update_idletasks()
            self._sync_numeric_settings()
            if self.signal is None or self.gray is None:
                self.load_image()
            if self.signal is None or self.gray is None:
                return
            if self.crop_start is not None and self.crop_current is not None:
                self.apply_crop()
                if self.crop_start is not None or self.crop_current is not None:
                    return
            outdir = Path(self.outdir_var.get()).expanduser()
            outdir.mkdir(parents=True, exist_ok=True)
            long = self.quantify_points()
            summary = self.make_summary(long)
            stem = Path(self.image_var.get()).stem or "gel"
            long_csv = outdir / f"{stem}_manual_long.csv"
            summary_csv = outdir / f"{stem}_manual_summary.csv"
            summary_xlsx = outdir / f"{stem}_manual_summary.xlsx"
            qc_png = outdir / f"{stem}_manual_qc.png"
            long.to_csv(long_csv, index=False)
            summary.to_csv(summary_csv, index=False)
            write_result_xlsx(summary_xlsx, {"Summary": summary, "Long": long})
            output_display = display_from_gray(
                self.gray,
                invert=bool(self.invert_output_var.get()),
                black_percent=float(self.black_percent_var.get()),
                white_percent=float(self.white_percent_var.get()),
            )
            self.write_qc_image(output_display, qc_png, long)
            self.save_session(default=True)
            messagebox.showinfo(
                "Done",
                "Quantification finished.\n\n"
                + "\n".join(str(path) for path in [long_csv, summary_csv, summary_xlsx, qc_png]),
            )
            self.status_var.set(f"Finished: {outdir}")
        except Exception as exc:
            messagebox.showerror("Run failed", str(exc))

    def quantify_points(self) -> pd.DataFrame:
        if self.signal is None or self.gray is None:
            raise RuntimeError("Load an image first.")
        height, width = self.signal.shape
        rows = []
        band_count = self._band_count()
        self.band_count_var.set(str(band_count))
        lane_half = int(self.lane_half_width_var.get())
        band_half = int(self.band_half_height_var.get())
        for lane_pos, (lane, x_center) in enumerate(zip(self.lanes, self.centers)):
            x1, x2 = clip_interval(x_center, lane_half, width)
            for band_idx in range(1, band_count + 1):
                band = str(band_idx)
                point = self.points.get((lane_pos, band))
                if point is None:
                    rows.append(
                        self._empty_row(lane, lane_pos, x_center, x1, x2, band)
                    )
                    continue
                peak_y = int(round(point.y))
                y1 = max(0, peak_y - band_half)
                y2 = min(height, peak_y + band_half + 1)
                roi_signal = self.signal[y1:y2, x1:x2]
                roi_gray = self.gray[y1:y2, x1:x2]
                bg = background_pixels(
                    self.signal,
                    x_center,
                    y1,
                    y2,
                    lane_half,
                    int(self.background_gap_var.get()),
                    int(self.background_width_var.get()),
                )
                background_median = float(np.nanmedian(bg)) if bg.size else 0.0
                corrected = np.maximum(roi_signal - background_median, 0.0)
                area = int(roi_signal.size)
                integrated = float(np.nansum(corrected)) if lane.include and area else 0.0
                rows.append(
                    {
                        "lane_pos": lane_pos + 1,
                        "lane_index": lane.lane_index,
                        "lane_label": lane.lane_label,
                        "label_top": lane.label_top,
                        "label_bottom": lane.label_bottom,
                        "include": lane.include,
                        "band": band,
                        "x_center": x_center,
                        "x1": x1,
                        "x2": x2,
                        "y1": y1,
                        "y2": y2,
                        "peak_y": peak_y,
                        "area_px": area,
                        "raw_gray_mean": float(np.nanmean(roi_gray)) if area else np.nan,
                        "raw_gray_integrated": float(np.nansum(roi_gray)) if area else np.nan,
                        "background_signal_median": background_median,
                        "background_corrected_integrated": integrated,
                        "mean_corrected_signal": integrated / area if area else np.nan,
                        "peak_profile_value": point.peak,
                        "note": lane.note,
                    }
                )
        return pd.DataFrame(rows)

    def _empty_row(
        self,
        lane: LaneSpec,
        lane_pos: int,
        x_center: int,
        x1: int,
        x2: int,
        band: str,
    ) -> dict[str, object]:
        return {
            "lane_pos": lane_pos + 1,
            "lane_index": lane.lane_index,
            "lane_label": lane.lane_label,
            "label_top": lane.label_top,
            "label_bottom": lane.label_bottom,
            "include": lane.include,
            "band": band,
            "x_center": x_center,
            "x1": x1,
            "x2": x2,
            "y1": np.nan,
            "y2": np.nan,
            "peak_y": np.nan,
            "area_px": 0,
            "raw_gray_mean": np.nan,
            "raw_gray_integrated": np.nan,
            "background_signal_median": np.nan,
            "background_corrected_integrated": 0.0,
            "mean_corrected_signal": np.nan,
            "peak_profile_value": 0.0,
            "note": lane.note,
        }

    def make_summary(self, long: pd.DataFrame) -> pd.DataFrame:
        index_cols = ["lane_pos", "lane_index", "lane_label", "label_top", "label_bottom", "include", "x_center", "note"]
        wide = long.pivot_table(
            index=index_cols,
            columns="band",
            values="background_corrected_integrated",
            aggfunc="first",
        ).reset_index()
        wide.columns.name = None
        band_names = [str(col) for col in wide.columns if col not in index_cols]
        rename = {band: f"band_{band}_integrated" for band in band_names}
        wide = wide.rename(columns=rename)
        integrated_cols = list(rename.values())
        if integrated_cols:
            wide["total_signal"] = wide[integrated_cols].sum(axis=1)
            for band, col in rename.items():
                wide[f"band_{band}_fraction"] = np.where(wide["total_signal"] > 0, wide[col] / wide["total_signal"], np.nan)
        return wide

    def write_qc_image(self, display: Image.Image, path: Path, long: pd.DataFrame) -> None:
        self._sync_numeric_settings()
        out = display.copy()
        draw = ImageDraw.Draw(out)
        self._draw_lane_labels(draw)
        marker_width = max(1, int(self.marker_width_var.get()))
        for _, row in long.iterrows():
            if float(row.get("peak_profile_value", 0.0)) <= 0:
                continue
            color = self.pil_color(self.color_for_band(str(row["band"])))
            peak_y = int(row["peak_y"])
            draw.rectangle(
                [
                    int(row["x1"]),
                    int(row["y1"]),
                    int(row["x2"]),
                    int(row["y2"]),
                ],
                outline=color,
                width=marker_width,
            )
            draw.line([(int(row["x1"]), peak_y), (int(row["x2"]), peak_y)], fill=color, width=marker_width)
        self._draw_band_legend(draw, long)
        out.save(path)

    def _draw_lane_labels(self, draw: ImageDraw.ImageDraw) -> None:
        label_y = self._int_value(self.label_y_var, DEFAULT_LABEL_Y, 0, 5000)
        top_size = self._int_value(self.top_font_size_var, DEFAULT_TOP_FONT_SIZE, 6, 40)
        bottom_size = self._int_value(self.bottom_font_size_var, DEFAULT_BOTTOM_FONT_SIZE, 6, 40)
        top_font = font(top_size, bold=True)
        bottom_font = font(bottom_size, bold=False)
        for idx, x_center in enumerate(self.centers):
            lane = self.lanes[idx]
            color = (0, 210, 230) if lane.include else (170, 170, 170)
            draw_centered(draw, x_center, label_y, lane.label_top, top_font, color)
            draw_centered(draw, x_center, label_y + top_size + 4, lane.label_bottom, bottom_font, color)
            y1 = label_y + top_size + bottom_size + 12
            y2 = y1 + 16
            draw.line([(x_center, y1), (x_center, y2)], fill=color, width=2)

    def _draw_band_legend(self, draw: ImageDraw.ImageDraw, long: pd.DataFrame) -> None:
        visible = long[pd.to_numeric(long["peak_profile_value"], errors="coerce").fillna(0.0) > 0]
        if visible.empty or not self.centers:
            return
        legend_font = font(self._int_value(self.band_font_size_var, DEFAULT_BAND_FONT_SIZE, 6, 40), bold=True)
        entries = self._band_legend_entries()
        if not entries:
            return
        lane_half = self._int_value(self.lane_half_width_var, 24, 3, 120)
        max_label_w = max(text_size(draw, f"band {band}", legend_font)[0] for band, _y in entries)
        left_x = max(4, min(self.centers) - lane_half - max_label_w - 42)
        for band, y_float in entries:
            y = int(round(y_float))
            color = self.pil_color(self.color_for_band(str(band)))
            label = f"band {band}"
            label_w, label_h = text_size(draw, label, legend_font)
            draw.text((left_x, y - int(label_h / 2)), label, font=legend_font, fill=color)
            draw.line([(left_x + label_w + 8, y), (left_x + label_w + 32, y)], fill=color, width=2)

    def pil_color(self, hex_color: str) -> tuple[int, int, int]:
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))

    def _session_data(self) -> dict[str, object]:
        self._sync_numeric_settings()
        return {
            "image": self.image_var.get(),
            "annotations": self.annotations_var.get(),
            "outdir": self.outdir_var.get(),
            "sheet": self.sheet_var.get(),
            "top_col": self.top_col_var.get(),
            "bottom_col": self.bottom_col_var.get(),
            "lane_col": self.lane_col_var.get(),
            "polarity": self.polarity_var.get(),
            "crop_box_original": self.crop_box_original,
            "settings": {
                "band_count": self._band_count(),
                "lane_half_width": int(self.lane_half_width_var.get()),
                "band_half_height": int(self.band_half_height_var.get()),
                "background_gap": int(self.background_gap_var.get()),
                "background_width": int(self.background_width_var.get()),
                "marker_width": int(self.marker_width_var.get()),
                "label_y": int(self.label_y_var.get()),
                "top_font_size": int(self.top_font_size_var.get()),
                "bottom_font_size": int(self.bottom_font_size_var.get()),
                "band_font_size": int(self.band_font_size_var.get()),
                "black_percent": float(self.black_percent_var.get()),
                "white_percent": float(self.white_percent_var.get()),
                "invert_output": bool(self.invert_output_var.get()),
            },
            "lanes": [asdict(lane) | {"center": center} for lane, center in zip(self.lanes, self.centers)],
            "points": [asdict(point) for point in self.points.values()],
            "legend_y_overrides": dict(self.legend_y_overrides),
        }

    def save_session(self, default: bool = False) -> None:
        try:
            outdir = Path(self.outdir_var.get()).expanduser()
            outdir.mkdir(parents=True, exist_ok=True)
            path = outdir / "gel_gui_session.json"
            if not default:
                chosen = filedialog.asksaveasfilename(
                    initialdir=str(outdir),
                    initialfile="gel_gui_session.json",
                    defaultextension=".json",
                    filetypes=[("JSON", "*.json"), ("All files", "*.*")],
                )
                if not chosen:
                    return
                path = Path(chosen)
            path.write_text(json.dumps(self._session_data(), indent=2), encoding="utf-8")
            self.status_var.set(f"Saved session: {path}")
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))

    def load_session(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=str(Path(self.outdir_var.get()).expanduser()),
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            self.image_var.set(str(data.get("image", self.image_var.get())))
            self.annotations_var.set(str(data.get("annotations", self.annotations_var.get())))
            self.outdir_var.set(str(data.get("outdir", self.outdir_var.get())))
            self.sheet_var.set(str(data.get("sheet", self.sheet_var.get())))
            self.top_col_var.set(str(data.get("top_col", self.top_col_var.get())))
            self.bottom_col_var.set(str(data.get("bottom_col", self.bottom_col_var.get())))
            self.lane_col_var.set(str(data.get("lane_col", self.lane_col_var.get())))
            polarity = str(data.get("polarity", self.polarity_var.get()))
            if polarity in {"bright", "dark"}:
                self.polarity_var.set(polarity)
            settings = data.get("settings", {})
            try:
                self.band_count_var.set(str(max(1, int(settings.get("band_count", self._band_count())))))
            except (TypeError, ValueError):
                self.band_count_var.set(str(self._band_count()))
            self.lane_half_width_var.set(int(settings.get("lane_half_width", self.lane_half_width_var.get())))
            self.band_half_height_var.set(int(settings.get("band_half_height", self.band_half_height_var.get())))
            self.background_gap_var.set(int(settings.get("background_gap", self.background_gap_var.get())))
            self.background_width_var.set(int(settings.get("background_width", self.background_width_var.get())))
            self.marker_width_var.set(int(settings.get("marker_width", self.marker_width_var.get())))
            self.label_y_var.set(int(settings.get("label_y", self.label_y_var.get())))
            self.top_font_size_var.set(int(settings.get("top_font_size", self.top_font_size_var.get())))
            self.bottom_font_size_var.set(int(settings.get("bottom_font_size", self.bottom_font_size_var.get())))
            self.band_font_size_var.set(int(settings.get("band_font_size", self.band_font_size_var.get())))
            self.black_percent_var.set(float(settings.get("black_percent", self.black_percent_var.get())))
            self.white_percent_var.set(float(settings.get("white_percent", self.white_percent_var.get())))
            self.invert_output_var.set(bool(settings.get("invert_output", self.invert_output_var.get())))
            self._refresh_band_buttons()

            original = read_gel_image(Path(self.image_var.get()).expanduser())
            self.original_gray = original
            crop_box = data.get("crop_box_original")
            self.crop_box_original = tuple(crop_box) if crop_box else None
            if self.crop_box_original:
                x1, y1, x2, y2 = [int(v) for v in self.crop_box_original]
                self._set_working_gray(original[y1:y2, x1:x2])
            else:
                self._set_working_gray(original)

            self.lanes = []
            self.centers = []
            for lane_data in data.get("lanes", []):
                self.lanes.append(
                    LaneSpec(
                        int(lane_data["lane_index"]),
                        str(lane_data["lane_label"]),
                        str(lane_data["label_top"]),
                        str(lane_data.get("label_bottom", "")),
                        bool(lane_data.get("include", True)),
                        None,
                        str(lane_data.get("note", "")),
                    )
                )
                self.centers.append(int(round(float(lane_data["center"]))))
            self.points.clear()
            for item in data.get("points", []):
                point = BandPoint(int(item["lane_pos"]), str(item["band"]), float(item["y"]), float(item.get("peak", 1.0)))
                self.points[(point.lane_pos, point.band)] = point
            self.legend_y_overrides = {
                str(band): float(y)
                for band, y in dict(data.get("legend_y_overrides", {})).items()
            }
            self._clear_history()
            self.status_var.set(f"Loaded session: {path}")
            self.redraw()
        except Exception as exc:
            messagebox.showerror("Load session failed", str(exc))


def main() -> None:
    root = tk.Tk()
    GelGuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
