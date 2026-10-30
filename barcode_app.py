"""Barcode Maker - pick items from the MASTER BARANG Excel export and print barcode stickers.

How it works:
  1. All items from the Master Barang file are listed.
  2. Tick the items to print and set the Print column to how many stickers each one needs.
  3. "Print Stickers" makes a Word file of stickers (built-in sticker sheet layout) and opens it in Word.

Each sticker shows the price code (e.g. AP0003/2SET), a Code 39 barcode of the Kode User and the Kode User.
"""

import copy
import csv
import io
import json
import math
import os
import queue
import shutil
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.opc.packuri import PackURI
from docx.oxml.ns import qn
from docx.package import ImageParts
from docx.parts.story import StoryPart
from docx.shared import Cm, Mm, Pt, Twips
from docx.table import _Cell
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# python-docx rescans every existing image and id each time a picture is added,
# which makes thousands of labels painfully slow. We only ever append, so counters work.
ImageParts._next_image_partname = lambda self, ext: PackURI(f"/word/media/image{len(self) + 1}.{ext}")
_original_next_id = StoryPart.next_id.fget


def _fast_next_id(self):
    n = self.__dict__.get("_next_id_counter") or _original_next_id(self)
    self.__dict__["_next_id_counter"] = n + 1
    return n


StoryPart.next_id = property(_fast_next_id)

# Where the app keeps its files (settings, SO drafts, the SO form, the error log): the folder of the
# .py file, or - when packed as an .exe - the folder of the .exe. BUNDLE_DIR is where files packed
# inside the .exe are unpacked while it runs.
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = getattr(sys, "_MEIPASS", APP_DIR)
else:
    APP_DIR = BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------- Code 39 --

# Each character is 9 elements: bar, space, bar, space, bar, space, bar, space, bar.
# "1" means a wide element, "0" a narrow one.
CODE39 = {
    "0": "000110100", "1": "100100001", "2": "001100001", "3": "101100000", "4": "000110001",
    "5": "100110000", "6": "001110000", "7": "000100101", "8": "100100100", "9": "001100100",
    "A": "100001001", "B": "001001001", "C": "101001000", "D": "000011001", "E": "100011000",
    "F": "001011000", "G": "000001101", "H": "100001100", "I": "001001100", "J": "000011100",
    "K": "100000011", "L": "001000011", "M": "101000010", "N": "000010011", "O": "100010010",
    "P": "001010010", "Q": "000000111", "R": "100000110", "S": "001000110", "T": "000010110",
    "U": "110000001", "V": "011000001", "W": "111000000", "X": "010010001", "Y": "110010000",
    "Z": "011010000", "-": "010000101", ".": "110000100", " ": "011000100", "$": "010101000",
    "/": "010100010", "+": "010001010", "%": "000101010", "*": "010010100",
}
START_STOP = "*"

# Full ASCII Code 39: characters outside the basic set are written as two-character pairs.
# Scanners must have "Code 39 Full ASCII" switched on to turn these pairs back into ( ) = etc.
FULL_ASCII = {
    "!": "/A", '"': "/B", "#": "/C", "&": "/F", "'": "/G", "(": "/H", ")": "/I", "*": "/J",
    ",": "/L", ":": "/Z", ";": "%F", "<": "%G", "=": "%H", ">": "%I", "?": "%J", "@": "%V",
    "[": "%K", "\\": "%L", "]": "%M", "^": "%N", "_": "%O", "`": "%W",
    "{": "%P", "|": "%Q", "}": "%R", "~": "%S",
}
FULL_ASCII.update({c: "+" + c.upper() for c in "abcdefghijklmnopqrstuvwxyz"})

WIDE_RATIO = 2.25  # wide element = 2.25 x narrow (Code 39 allows 2.0 - 3.0)
QUIET = 10         # blank narrow-widths on each side so scanners can find the edges


def encode39(text):
    """Return the Code 39 characters to draw for text, including the * start/stop characters."""
    if not text:
        raise ValueError("kode kosong")
    out = []
    for ch in text:
        if ch in CODE39 and ch != "*":
            out.append(ch)
        elif ch in FULL_ASCII:
            out.append(FULL_ASCII[ch])
        else:
            raise ValueError(f"karakter tidak didukung {ch!r}")
    return START_STOP + "".join(out) + START_STOP


def barcode_image(text, narrow_px=4, height_px=90, quiet=QUIET):
    """Draw a Code 39 barcode as a black/white image.

    Returns (image, width_in_narrow_units); a narrow bar printed at X mm makes the barcode
    width_in_narrow_units * X mm wide.
    """
    wide_px = round(narrow_px * WIDE_RATIO)
    elements = []  # (is_bar, width_px)
    for n, ch in enumerate(encode39(text)):
        if n:
            elements.append((False, narrow_px))  # narrow gap between characters
        for k, wide in enumerate(CODE39[ch]):
            elements.append((k % 2 == 0, wide_px if wide == "1" else narrow_px))
    width_px = sum(w for _, w in elements) + 2 * quiet * narrow_px
    img = Image.new("1", (width_px, height_px), 1)
    draw = ImageDraw.Draw(img)
    x = quiet * narrow_px
    for is_bar, w in elements:
        if is_bar:
            draw.rectangle([x, 0, x + w - 1, height_px - 1], fill=0)
        x += w
    return img, width_px / narrow_px


# ------------------------------------------------------------- Price code --

# Sat -> code used in the price code. A Sat that is not listed here uses its first 3 letters.
SAT_CODES = {
    "PCS": "BJ",
    "SET": "SET",
    "PAK": "PK",
    "BALL": "BALL",
    "BOX": "BX",
}


def kode_harga(harga, sat):
    """Price code for a sticker: thousands / the 3 digits below 1000 without their trailing zeros + sat
    code. Rp 3.200 SET -> AP0003/2SET, Rp 2.050 -> AP0002/05SET, Rp 1.001 -> AP0001/001SET,
    Rp 2.000 -> AP0002/0SET."""
    sat = sat.strip().rstrip(".").upper()
    sat_code = SAT_CODES.get(sat, sat[:3]) if sat not in ("", "-") else ""
    rest = f"{harga % 1000:03d}".rstrip("0") or "0"
    return f"AP{harga // 1000:04d}/{rest}{sat_code}"


# ------------------------------------------------------------ Excel input --

COL_KODE, COL_NAMA, COL_HARGA, COL_SAT = "Kode User", "Nama Barang", "Hrg. Jual 1", "Sat."


def read_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        import xlrd
        sheet = xlrd.open_workbook(path).sheet_by_index(0)
        return [sheet.row_values(r) for r in range(sheet.nrows)]
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        return [["" if v is None else v for v in row] for row in ws.iter_rows(values_only=True)]
    raise ValueError("Silakan pilih file Excel (.xls atau .xlsx).")


def load_items(path):
    """Return (items, skipped_count). Each item is (kode, nama, harga, sat)."""
    rows = read_rows(path)
    header_row = None
    for r, row in enumerate(rows[:30]):
        cells = [str(c).strip() for c in row]
        if all(c in cells for c in (COL_KODE, COL_NAMA, COL_HARGA, COL_SAT)):
            header_row = r
            # index() finds the first "Sat." column, which is the unit for Hrg. Jual
            i_kode, i_nama, i_harga, i_sat = (cells.index(c) for c in (COL_KODE, COL_NAMA, COL_HARGA, COL_SAT))
            break
    if header_row is None:
        raise ValueError(f"Kolom '{COL_KODE}', '{COL_NAMA}', '{COL_HARGA}' dan '{COL_SAT}' tidak ditemukan.")

    items, skipped, seen = [], 0, set()
    for row in rows[header_row + 1:]:
        if len(row) <= max(i_kode, i_nama, i_harga, i_sat):
            continue
        kode = str(row[i_kode]).strip()
        nama = str(row[i_nama]).strip()
        if kode == COL_KODE:  # repeated header row
            continue
        if not kode:
            if nama:
                skipped += 1  # item without a Kode User can't get a barcode
            continue
        if kode in seen:
            continue
        try:
            encode39(kode)
        except ValueError:
            skipped += 1
            continue
        seen.add(kode)
        try:
            harga = int(round(float(row[i_harga])))
        except (TypeError, ValueError):
            harga = 0
        items.append((kode, nama, harga, str(row[i_sat]).strip()))
    return items, skipped


def angka(n):
    """Number with Indonesian thousands separator: 14422 -> 14.422."""
    return f"{n:,}".replace(",", ".")


def rupiah(n):
    return "Rp " + angka(n)


# ------------------------------------------------- Export all (A4 labels) --

def write_csv(items, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig so Excel reads it right
        w = csv.writer(f)
        w.writerow(["Kode Barang", "Nama Barang", "Harga", "Sat", "Kode Harga"])
        w.writerows((kode, nama, harga, sat, kode_harga(harga, sat)) for kode, nama, harga, sat in items)


LABEL_COLS = 3
MODULE_MM = 0.25   # printed width of the thinnest bar
MAX_BARCODE_MM = 60
LABELS_PER_FILE = 1500


def _cell_margins(table, cm):
    tbl_pr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(int(cm * 567)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.append(mar)


def _no_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    tr_pr.append(el)


def write_docx(items, path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)  # A4
    sec.left_margin = sec.right_margin = Cm(1)
    sec.top_margin = sec.bottom_margin = Cm(1)
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(8)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    rows = (len(items) + LABEL_COLS - 1) // LABEL_COLS
    table = doc.add_table(rows=rows, cols=LABEL_COLS)
    table.style = "Table Grid"  # thin borders act as cutting guides
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _cell_margins(table, 0.15)

    cells = []
    for row in table.rows:  # collect cells once; indexing table.rows repeatedly is very slow
        _no_split(row)
        cells.extend(row.cells)

    for cell, (kode, nama, harga, _sat) in zip(cells, items):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        r = p.add_run(nama[:60])
        r.font.size = Pt(7)

        img, units = barcode_image(kode)
        buf = io.BytesIO()
        img.save(buf, "PNG")
        buf.seek(0)
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.add_run().add_picture(buf, width=Mm(min(units * MODULE_MM, MAX_BARCODE_MM)), height=Mm(10))

        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(kode)
        r.font.size = Pt(8)
        r.font.name = "Consolas"

        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(rupiah(harga))
        r.font.size = Pt(10)
        r.bold = True

    doc.save(path)


# ------------------------------------------------------- Sticker printing --

STICKER_BARCODE_MM = 8    # barcode height on a sticker
TWIPS_PER_MM = 56.7
SLOT, OFF_PAGE, BLANK = "slot", "off", "blank"

# Sticker sheet layouts, all sizes in mm. A layout is: page size, where the first sticker starts
# (margin + optional blank gap strips), how many stickers across/down, and the sticker size.
# sticker_w can be a list to give each column its own width.
# The sticker paper, as measured: 7 rows of boxes starting at the top edge - the top row 21 mm high,
# the others 19 mm; the first box is 29 mm wide, the other 7 are 25 mm; 11 mm of empty paper on the
# right and 3 mm at the bottom.
PAPER_COLS = [29.0] + [25.0] * 7   # box widths left to right (mm), starting at the left edge
PAPER_ROW_HEIGHTS = [21.0] + [19.0] * 6   # box heights top to bottom (mm)
PAPER_ROWS = len(PAPER_ROW_HEIGHTS)
PAPER_ROW_H = min(PAPER_ROW_HEIGHTS)      # the smallest box decides the barcode size
PAPER_RIGHT_EMPTY = 11.0           # empty paper right of the last box
PAPER_BOTTOM_EMPTY = 3.0           # empty paper below the last row
PAPER_LEFT_STRIP = 0.0             # move everything right (mm) if the print needs nudging
PAPER_TOP_STRIP = 0.0              # move everything down (mm) if the print needs nudging
PAPER_TEXT_RIGHT = 0.0             # move the text and barcode inside every box right (mm); boxes stay put
# Where the print sits inside the boxes of a row: "center" (default), "top" or "bottom". Row 1 starts at
# the paper's top edge, which printers can't print, so its print sits at the bottom of its 21 mm box.
PAPER_ROW_ALIGN = {1: "bottom"}
# The paper goes into the printer with its left edge (no empty strip) first: print portrait, with the
# whole layout turned a quarter turn. False = print the layout as it is (landscape).
PRINT_PORTRAIT = True
# Portrait only: move the print (text and barcode, not the box) of one row to the right on the page,
# in mm: {row: mm}, row 1 = the rightmost column. Negative = to the left. Row 7 is the leftmost
# column, next to the paper edge printers can't reach, so its print is moved inward.
PORTRAIT_NUDGE_RIGHT = {1: 2.0, 7: 2.0}
# Make the barcodes smaller than the box allows (mm). Shorter is safe for scanning; narrower makes the
# bars thinner, so long codes scan less easily (the app warns about those when printing). Together with
# the smaller text (STICKER_FONT_PT 8 instead of 9) this frees about 1.5 mm of height in every box.
BARCODE_SHORTER = 0.75
BARCODE_NARROWER = 0.5     # up to 0.5 keeps the long codes scannable; 0.6+ makes ~1000 more codes thin
STICKER_PAGE = (sum(PAPER_COLS) + PAPER_RIGHT_EMPTY, sum(PAPER_ROW_HEIGHTS) + PAPER_BOTTOM_EMPTY)  # 215 x 138
# (row height, space at the top, space at the bottom) in mm, row 1 first (measured: the rows come out
# exactly this far apart). To nudge one row: space at the top moves that row's print down; space at the
# bottom (Word adds it below the row) pushes all rows below it down.
STICKER_ROWS = [(height, 0, 0) for height in PAPER_ROW_HEIGHTS]
LAYOUTS = {
    "Asli: 8 x 7 stiker": {
        "page_w": STICKER_PAGE[0], "page_h": STICKER_PAGE[1], "margin_left": PAPER_LEFT_STRIP,
        "margin_top": PAPER_TOP_STRIP, "margin_bottom": 0, "gap_left": 0, "gap_top": 0,
        "cols": len(PAPER_COLS), "rows": PAPER_ROWS, "sticker_w": PAPER_COLS, "sticker_h": PAPER_ROW_H,
        "row_plan": STICKER_ROWS, "text_shift_x": PAPER_TEXT_RIGHT, "row_align": PAPER_ROW_ALIGN,
        "barcode_shorter": BARCODE_SHORTER, "barcode_narrower": BARCODE_NARROWER, "portrait": PRINT_PORTRAIT,
        "portrait_nudge": PORTRAIT_NUDGE_RIGHT,
    },
    "Lebar: 6 x 7 stiker": {  # same paper and rows, 6 wider stickers across (cut by hand)
        "page_w": STICKER_PAGE[0], "page_h": STICKER_PAGE[1], "margin_left": PAPER_LEFT_STRIP,
        "margin_top": PAPER_TOP_STRIP, "margin_bottom": 0, "gap_left": 0, "gap_top": 0,
        "cols": 6, "rows": PAPER_ROWS, "sticker_w": sum(PAPER_COLS) / 6, "sticker_h": PAPER_ROW_H,
        "row_plan": STICKER_ROWS, "text_shift_x": PAPER_TEXT_RIGHT, "row_align": PAPER_ROW_ALIGN,
        "barcode_shorter": BARCODE_SHORTER, "barcode_narrower": BARCODE_NARROWER, "portrait": PRINT_PORTRAIT,
        "portrait_nudge": PORTRAIT_NUDGE_RIGHT,
    },
}
# The paper sizes as one setting, so the "Ukuran Kertas" window (gear button) can change them;
# the numbers above are the defaults.
DEFAULT_PAPER = {
    "cols": list(PAPER_COLS), "rows": list(PAPER_ROW_HEIGHTS),
    "right_empty": PAPER_RIGHT_EMPTY, "bottom_empty": PAPER_BOTTOM_EMPTY,
    "left": PAPER_LEFT_STRIP, "top": PAPER_TOP_STRIP,
}
PAPER = dict(DEFAULT_PAPER)


def check_paper(paper):
    """A clean copy of a paper description, or ValueError with a readable message."""
    try:
        clean = {"cols": [float(v) for v in paper["cols"]], "rows": [float(v) for v in paper["rows"]],
                 "right_empty": float(paper["right_empty"]), "bottom_empty": float(paper["bottom_empty"]),
                 "left": float(paper.get("left", 0)), "top": float(paper.get("top", 0))}
    except (KeyError, TypeError, ValueError):
        raise ValueError("Semua ukuran harus angka (mm).")
    if not clean["cols"] or not clean["rows"]:
        raise ValueError("Harus ada kotak.")
    if any(v < 10 for v in clean["cols"]) or any(v < 12 for v in clean["rows"]):
        raise ValueError("Lebar kotak minimal 10 mm dan tinggi kotak minimal 12 mm.")
    if min(clean["right_empty"], clean["bottom_empty"], clean["left"], clean["top"]) < 0:
        raise ValueError("Ruang kosong tidak boleh negatif.")
    return clean


def apply_paper(paper):
    """Use these paper sizes for the two built-in layouts (updated in place, so the app sees it)."""
    global PAPER
    paper = check_paper(paper)
    page_w = paper["left"] + sum(paper["cols"]) + paper["right_empty"]
    page_h = paper["top"] + sum(paper["rows"]) + paper["bottom_empty"]
    common = {"page_w": page_w, "page_h": page_h, "margin_left": paper["left"], "margin_top": paper["top"],
              "rows": len(paper["rows"]), "sticker_h": min(paper["rows"]),
              "row_plan": [(h, 0, 0) for h in paper["rows"]]}
    LAYOUTS[_ASLI].update(common, cols=len(paper["cols"]), sticker_w=list(paper["cols"]))
    LAYOUTS[_LEBAR].update(common, cols=6, sticker_w=sum(paper["cols"]) / 6)
    PAPER = paper


# names used by earlier versions of the app, so a saved choice keeps working
_ASLI, _LEBAR = list(LAYOUTS)
OLD_LAYOUT_NAMES = {
    "Asli: 8 x 7 stiker (kertas 21 x 13 cm)": _ASLI,
    "Lebar: 6 x 7 stiker (kertas 21 x 13 cm)": _LEBAR,
    "Original: 8 x 7 stickers, 25 x 19 mm": _ASLI,
    "Asli: 8 x 7 stiker, 25 x 19 mm": _ASLI,
    "Asli: 8 x 7 stiker, 25 x 18,3 mm": _ASLI,
    "Wide: 6 x 7 stickers, 34.5 x 19 mm": _LEBAR,
    "Lebar: 6 x 7 stiker, 34,5 x 19 mm": _LEBAR,
    "Lebar: 6 x 7 stiker, 33,3 x 18,3 mm": _LEBAR,
}
DEFAULT_LAYOUT = next(iter(LAYOUTS))
CUSTOM = "Kustom..."

STICKER_FONT_PT = 8      # text on the stickers (pt)
STICKER_PADDING = 74      # space left and right inside a sticker (1.3 mm), also the barcode's quiet zone
MIN_BAR_MM = 0.12         # narrow bars thinner than this may not scan; such codes are reported


def _tw(mm):
    return Twips(round(mm * TWIPS_PER_MM))


def sticker_widths(layout):
    w = layout["sticker_w"]
    return list(w) if isinstance(w, (list, tuple)) else [w] * layout["cols"]


def stickers_per_sheet(layout):
    return layout["cols"] * layout["rows"]


def side_padding(layout):
    """Space (twips) left and right inside every sticker. text_shift_x (mm) moves the print right by
    taking that much from the right side and adding it on the left, so the barcode keeps its width."""
    shift = round(layout.get("text_shift_x", 0) * TWIPS_PER_MM)
    left, right = STICKER_PADDING + shift, STICKER_PADDING - shift
    if right < 0:  # shifting further than the padding: the extra comes off the usable width
        left, right = left - right, 0
    if left < 0:
        left, right = 0, right - left
    return left, right


def barcode_space_mm(layout):
    """Width available for the barcode on the narrowest sticker."""
    return (min(sticker_widths(layout)) - (sum(side_padding(layout)) + 30) / TWIPS_PER_MM
            - layout.get("barcode_narrower", 0))


def barcode_height_mm(layout):
    # leave room for the two text lines, then make it as much shorter as the layout asks
    return max(4, min(12, layout["sticker_h"] - 11) - layout.get("barcode_shorter", 0))


BORDER_MM = 0.5 / 72 * 25.4  # each dotted line between sticker rows is 0.5 pt


def row_shifts(layout):
    """How far (mm, + = down) each sticker row is moved: one number per row, top row first."""
    shifts = {int(k): float(v) for k, v in layout.get("row_shift", {}).items()}  # keys may be text (JSON)
    return [shifts.get(r, 0.0) for r in range(1, layout["rows"] + 1)]


def row_plan(layout):
    """Height and top/bottom padding (mm) of each sticker row, so that each row's print lands
    `row_shifts` lower than normal while the rows stay one solid table.

    A row is made taller or shorter by the difference to the next row's shift, which moves every row
    below it; padding on that row puts its own (centred) print back where it belongs.
    """
    # Word quirk (measured): in a row of exact height, top padding is inside the row, but bottom padding
    # is added below it. So "taller row with padding at the bottom" = normal height + bottom padding.
    if layout.get("row_plan"):  # rows given one by one (the sticker paper layouts)
        return [tuple(float(v) for v in row) for row in layout["row_plan"]]
    s, h = row_shifts(layout), layout["sticker_h"]
    plan = []
    for r in range(layout["rows"]):
        delta = s[r + 1] - s[r] if r + 1 < layout["rows"] else 0.0
        if delta >= 0:
            plan.append([h, 0.0, delta])            # the bottom padding pushes the rows below down
        else:
            plan.append([h + delta, -delta, 0.0])   # shorter row, print kept in place by top padding
    # Word adds the dotted lines between rows to the table height; if the rows now run past the page,
    # take the excess off the bottom of the last row (and add it as top padding to keep its print in place)
    lines = (layout["rows"] + 1) * BORDER_MM
    bottom = layout["margin_top"] + layout["gap_top"] + s[0] + sum(p[0] + p[2] for p in plan) + lines
    excess = bottom - (layout["page_h"] - 1.0)  # room for the tiny paragraph Word needs after the table
    if excess > 0 and s[-1] > 0:
        cut = min(excess, s[-1])
        plan[-1][0] -= cut
        plan[-1][1] += cut
    return [tuple(p) for p in plan]


def check_layout(layout):
    """Raise ValueError with a readable message if the stickers don't fit on the page."""
    across = layout["margin_left"] + layout["gap_left"] + sum(sticker_widths(layout))
    down = (layout["margin_top"] + layout["gap_top"] + row_shifts(layout)[0]
            + sum(height + pad_bottom for height, _, pad_bottom in row_plan(layout)))
    if layout["cols"] < 1 or layout["rows"] < 1:
        raise ValueError("Minimal harus ada 1 kolom dan 1 baris.")
    if layout["sticker_h"] < 12:
        raise ValueError("Tinggi stiker minimal 12 mm agar teks dan barcode muat.")
    if across > layout["page_w"] + 0.01:
        raise ValueError(f"Total lebar stiker {across:.1f} mm, tetapi lebar halaman hanya "
                         f"{layout['page_w']:.1f} mm.")
    if down > layout["page_h"] + 0.01:
        raise ValueError(f"Total tinggi stiker {down:.1f} mm, tetapi tinggi halaman hanya "
                         f"{layout['page_h']:.1f} mm.")


def build_sticker_template(layout):
    """Make a one-sheet sticker layout (Word file in memory) with placeholder text in every sticker."""
    check_layout(layout)
    widths = ([layout["gap_left"]] if layout["gap_left"] else []) + sticker_widths(layout)
    plan = row_plan(layout)
    row_align = {int(k): v for k, v in layout.get("row_align", {}).items()}  # keys may be text (JSON)
    heights = ([layout["gap_top"]] if layout["gap_top"] else []) + [height for height, _, _ in plan]
    first_col, first_row = len(widths) - layout["cols"], len(heights) - layout["rows"]
    top_margin = layout["margin_top"] + row_shifts(layout)[0]  # the first row's shift moves the table

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = _tw(layout["page_w"]), _tw(layout["page_h"])
    sec.left_margin, sec.top_margin = _tw(layout["margin_left"]), _tw(top_margin)
    sec.bottom_margin = _tw(layout.get("margin_bottom", 0))
    sec.right_margin = _tw(max(0, layout["page_w"] - layout["margin_left"] - sum(widths) - 0.5))
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(STICKER_FONT_PT)
    normal.paragraph_format.space_before = normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0

    table = doc.add_table(rows=len(heights), cols=len(widths))
    table.style = "Table Grid"
    table.autofit = False
    borders = OxmlElement("w:tblBorders")  # light dotted lines to cut along
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        for k, v in (("val", "dotted"), ("sz", "4"), ("space", "0"), ("color", "BFBFBF")):
            el.set(qn(f"w:{k}"), v)
        borders.append(el)
    table._tbl.tblPr.append(borders)
    padding = OxmlElement("w:tblCellMar")
    pad_left, pad_right = side_padding(layout)
    # Word lines a table up by the left padding of its first cell, so more left padding would move the
    # boxes instead of the print. Indenting the table by the extra padding keeps the boxes in place.
    if pad_left != STICKER_PADDING:
        ind = OxmlElement("w:tblInd")
        ind.set(qn("w:w"), str(pad_left - STICKER_PADDING))
        ind.set(qn("w:type"), "dxa")
        tbl_pr = table._tbl.tblPr
        tbl_pr.insert(list(tbl_pr).index(tbl_pr.find(qn("w:tblW"))) + 1, ind)  # tblInd comes after tblW
    for side, w in (("top", 0), ("left", STICKER_PADDING), ("bottom", 0), ("right", STICKER_PADDING)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(w))
        el.set(qn("w:type"), "dxa")
        padding.append(el)
    table._tbl.tblPr.append(padding)

    for c, w in enumerate(widths):
        table.columns[c].width = _tw(w)
    for r, row in enumerate(table.rows):
        row.height = _tw(heights[r])
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        _, pad_top, pad_bottom = plan[r - first_row] if r >= first_row else (0, 0, 0)
        valign = {"top": WD_ALIGN_VERTICAL.TOP, "bottom": WD_ALIGN_VERTICAL.BOTTOM}.get(
            row_align.get(r - first_row + 1, "center"), WD_ALIGN_VERTICAL.CENTER)
        for c, cell in enumerate(row.cells):
            cell.width = _tw(widths[c])
            cell.vertical_alignment = valign
            if r < first_row or c < first_col:
                continue  # gap strips stay empty
            sides = [(side, round(mm * TWIPS_PER_MM)) for side, mm in (("top", pad_top), ("bottom", pad_bottom)) if mm]
            if (pad_left, pad_right) != (STICKER_PADDING, STICKER_PADDING):  # print moved sideways
                sides += [("left", pad_left), ("right", pad_right)]
            if sides:  # padding of this sticker cell: top/bottom from row_plan, left/right from side_padding
                order = {"top": 0, "left": 1, "bottom": 2, "right": 3}  # the order Word expects
                mar = OxmlElement("w:tcMar")
                for side, tw in sorted(sides, key=lambda item: order[item[0]]):
                    el = OxmlElement(f"w:{side}")
                    el.set(qn("w:w"), str(tw))
                    el.set(qn("w:type"), "dxa")
                    mar.append(el)
                tc_pr = cell._tc.get_or_add_tcPr()
                tc_pr.insert(list(tc_pr).index(tc_pr.find(qn("w:vAlign"))), mar)  # tcMar goes before vAlign
            for i, text in enumerate(("AP0000/0", "KODE")):
                p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                run.bold = True
                run.font.size = Pt(STICKER_FONT_PT)
    doc.add_paragraph()  # Word wants a paragraph after a table

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def _p_text(p):
    return p.xpath("string(.)").strip()


def read_template(path):
    """Open a sticker layout (Word file path or file object) and work out which table cells are stickers.

    Returns (doc, table, layout). layout[row][cell] is (kind, usable_width_mm): a SLOT is a
    sticker (a cell with text in the template), OFF_PAGE cells lie outside the printable page.
    """
    doc = Document(path)
    if not doc.tables:
        raise ValueError("File Word stiker tidak berisi tabel stiker.")
    table = doc.tables[0]
    tbl = table._tbl
    sec = doc.sections[0]
    text_width = (sec.page_width - sec.left_margin - sec.right_margin) / 635  # EMU -> twips
    grid = [int(gc.get(qn("w:w"))) for gc in tbl.tblGrid.findall(qn("w:gridCol"))]
    margin = {"left": 108, "right": 108}  # Word's default cell padding (0.19 cm)
    cell_mar = tbl.tblPr.find(qn("w:tblCellMar"))
    if cell_mar is not None:
        for side, alt in (("left", "start"), ("right", "end")):
            el = cell_mar.find(qn(f"w:{side}"))
            if el is None:
                el = cell_mar.find(qn(f"w:{alt}"))
            if el is not None:
                margin[side] = int(el.get(qn("w:w")))

    layout = []
    for tr in tbl.tr_lst:
        col, row = 0, []
        before = tr.trPr.find(qn("w:gridBefore")) if tr.trPr is not None else None
        if before is not None:
            col = int(before.get(qn("w:val")))
        for tc in tr.tc_lst:
            x, width = sum(grid[:col]), sum(grid[col:col + tc.grid_span])
            col += tc.grid_span
            usable_mm = (width - margin["left"] - margin["right"] - 30) / TWIPS_PER_MM
            if x >= text_width:
                row.append((OFF_PAGE, usable_mm))
            elif any(_p_text(p) for p in tc.p_lst):
                row.append((SLOT, usable_mm))
            else:
                row.append((BLANK, usable_mm))
        layout.append(row)
    if not any(kind == SLOT for row in layout for kind, _ in row):
        raise ValueError("File Word stiker tidak berisi stiker bertulisan untuk dipakai sebagai tata letak.")
    return doc, table, layout


def _strip_paragraph(p):
    """Remove everything from a paragraph except its formatting."""
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def _set_text(p, text, usable_mm):
    """Replace a paragraph's text, keeping the first run's formatting and shrinking long text to fit."""
    runs = p.findall(qn("w:r"))
    rpr = runs[0].find(qn("w:rPr")) if runs else None
    _strip_paragraph(p)
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    p.append(r)
    run = Run(r, None)
    run.text = text
    size = run.font.size.pt if run.font.size else 11
    # rough width of bold text: about 0.2 mm per character per point of font size
    fit = usable_mm / (0.2 * max(len(text), 1))
    if fit < size:
        run.font.size = Pt(max(5, math.floor(fit * 2) / 2))


def _fill_slot(tc, table, sticker, usable_mm, barcode_mm=STICKER_BARCODE_MM, narrower=0.0):
    """Write one sticker into a table cell. Returns the printed width of its narrow bars in mm."""
    text_ps = [p for p in tc.p_lst if _p_text(p)]
    top_tpl, code_tpl = text_ps[0], text_ps[-1]
    for p in tc.p_lst:
        tc.remove(p)

    if sticker is None:  # no more items: leave this sticker blank
        blank = copy.deepcopy(top_tpl)
        _strip_paragraph(blank)
        tc.append(blank)
        return None

    top_text, code = sticker
    if top_text:
        p = copy.deepcopy(top_tpl)
        _set_text(p, top_text, usable_mm)
        tc.append(p)

    # quiet=0: the sticker's own padding is the blank margin around the barcode
    img, units = barcode_image(code, narrow_px=8, height_px=120, quiet=0)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    width_mm = min(usable_mm - narrower, units * 0.33)  # narrower: keep the barcode that much under the box width
    p = copy.deepcopy(code_tpl)
    _strip_paragraph(p)
    tc.append(p)
    para = Paragraph(p, _Cell(tc, table))
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after = Pt(1)
    para.add_run().add_picture(buf, width=Mm(width_mm), height=Mm(barcode_mm))

    p = copy.deepcopy(code_tpl)
    _set_text(p, code, usable_mm)
    tc.append(p)
    return width_mm / units


def _clear_cell(tc):
    ps = tc.p_lst
    for p in ps[1:]:
        tc.remove(p)
    _strip_paragraph(ps[0])


def make_sticker_doc_portrait(dst, stickers, layout):
    """Like make_sticker_doc, but for paper that goes into the printer with its left edge first: the
    page is turned to portrait and the whole sticker layout (boxes, text and barcodes) is rotated a
    quarter turn clockwise, so it lands on the sheet exactly where the landscape layout would.

    Landscape position (x right, y down) -> portrait position (page height - y, x): the sheet's left
    edge becomes the top of the page, its top edge the right side. So the landscape columns become
    the portrait rows (top to bottom), and the landscape rows the portrait columns (right to left).
    """
    check_layout(layout)
    widths = sticker_widths(layout)             # landscape box widths -> portrait row heights
    plan = row_plan(layout)                     # landscape row heights -> portrait column widths
    heights = [height for height, _, _ in plan]
    row_align = {int(k): v for k, v in layout.get("row_align", {}).items()}
    nudges = {int(k): float(v) for k, v in layout.get("portrait_nudge", {}).items()}
    rows, cols = layout["rows"], layout["cols"]
    per_sheet = rows * cols
    sheets = max(1, math.ceil(len(stickers) / per_sheet))
    barcode_mm = barcode_height_mm(layout)
    narrower = layout.get("barcode_narrower", 0.0)
    padding_mm = (2 * STICKER_PADDING + 30) / TWIPS_PER_MM   # the same room the landscape layout leaves

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = _tw(layout["page_h"]), _tw(layout["page_w"])
    # landscape top margin -> portrait right side; landscape left margin -> portrait top
    sec.top_margin, sec.bottom_margin, sec.right_margin = _tw(layout["margin_left"]), 0, 0
    sec.left_margin = _tw(max(0.0, layout["page_h"] - layout["margin_top"] - sum(heights)))
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(STICKER_FONT_PT)
    normal.paragraph_format.space_before = normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0

    table = doc.add_table(rows=0, cols=rows)
    table.style = "Table Grid"
    table.autofit = False
    borders = OxmlElement("w:tblBorders")  # light dotted lines to cut along
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        for k, v in (("val", "dotted"), ("sz", "4"), ("space", "0"), ("color", "BFBFBF")):
            el.set(qn(f"w:{k}"), v)
        borders.append(el)
    table._tbl.tblPr.append(borders)
    padding = OxmlElement("w:tblCellMar")   # no cell padding: the room is left by sizing the print instead
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        padding.append(el)
    table._tbl.tblPr.append(padding)
    col_widths = heights[::-1]                  # portrait column j = landscape row (rows - 1 - j)
    for j, w in enumerate(col_widths):
        table.columns[j].width = _tw(w)

    thin_codes = {}
    for sheet in range(sheets):
        for c in range(cols):                   # portrait row = landscape column c
            row = table.add_row()
            row.height = _tw(widths[c])
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            usable = widths[c] - padding_mm     # the print runs along the box's length
            for j, cell in enumerate(row.cells):
                r = rows - 1 - j                # landscape row of this cell
                cell.width = _tw(col_widths[j])
                align = row_align.get(r + 1, "center")
                cell.vertical_alignment = {"top": WD_ALIGN_VERTICAL.TOP,
                                           "bottom": WD_ALIGN_VERTICAL.BOTTOM}.get(align, WD_ALIGN_VERTICAL.CENTER)
                tc_pr = cell._tc.get_or_add_tcPr()
                direction = OxmlElement("w:textDirection")
                direction.set(qn("w:val"), "tbRl")   # text turned a quarter turn clockwise
                tc_pr.insert(list(tc_pr).index(tc_pr.find(qn("w:vAlign"))), direction)
                # Nudge: with vertical text, paragraph spacing lies sideways on the page. Extra space after
                # the last line (the Kode User, on the left) pushes the print right; extra space before the
                # first line (on the right) pushes it left. The print moves by the full space when it sits
                # against that side (row aligned to it), by half when it is centred.
                nudge = nudges.get(r + 1, 0.0)
                against = (align == "bottom" and nudge > 0) or (align == "top" and nudge < 0)
                nudge_space = Mm(abs(nudge) * (1 if against else 2)) if nudge else None

                index = sheet * per_sheet + r * cols + c
                sticker = stickers[index] if index < len(stickers) else None
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if sticker is None:
                    continue
                top_text, code = sticker
                if top_text:  # in vertical text the first paragraph is on the right = the sticker's top
                    _add_fitted_run(p, top_text, usable)
                    if nudge < 0:
                        p.paragraph_format.space_before = nudge_space
                    p = cell.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img, units = barcode_image(code, narrow_px=8, height_px=120, quiet=0)
                length_mm = min(usable - narrower, units * 0.33)
                buf = io.BytesIO()
                img.rotate(-90, expand=True).save(buf, "PNG")   # Word doesn't turn pictures, so turn it here
                buf.seek(0)
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.add_run().add_picture(buf, width=Mm(barcode_mm), height=Mm(length_mm))
                p = cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _add_fitted_run(p, code, usable)
                if nudge > 0:
                    p.paragraph_format.space_after = nudge_space
                if length_mm / units < MIN_BAR_MM:
                    thin_codes[code] = None
            if sheet and c == 0:  # every sheet starts on a new page
                row.cells[0].paragraphs[0].paragraph_format.page_break_before = True

    end = doc.add_paragraph()  # Word wants a paragraph after a table; keep it tiny so it fits on the page
    fmt = end.paragraph_format
    fmt.space_before = fmt.space_after = Pt(0)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(1)
    doc.save(dst)
    return sheets, list(thin_codes)


def _add_fitted_run(p, text, usable_mm):
    """Bold sticker text, shrunk if it would be longer than the room it has (same rule as _set_text)."""
    run = p.add_run(text)
    run.bold = True
    size = STICKER_FONT_PT
    fit = usable_mm / (0.2 * max(len(text), 1))
    run.font.size = Pt(max(5, math.floor(fit * 2) / 2) if fit < size else size)


def make_sticker_doc(dst, stickers, layout=None, template=None):
    """Fill a sticker sheet layout with stickers [(top_text, kode), ...], adding sheets as needed.

    Uses `layout` (one of LAYOUTS, or a custom one) unless a template Word file is given.
    Returns (sheets, thin_codes): thin_codes are codes so long that their bars may be too thin to scan.
    """
    layout = layout or LAYOUTS[DEFAULT_LAYOUT]
    if layout.get("portrait") and not template:
        return make_sticker_doc_portrait(dst, stickers, layout)
    if template:
        doc, table, cells = read_template(template)
        barcode_mm, narrower = STICKER_BARCODE_MM, 0.0
    else:
        doc, table, cells = read_template(build_sticker_template(layout))
        barcode_mm, narrower = barcode_height_mm(layout), layout.get("barcode_narrower", 0.0)
    tbl = table._tbl
    per_sheet = sum(kind == SLOT for row in cells for kind, _ in row)
    sheets = max(1, math.ceil(len(stickers) / per_sheet))
    template_rows = list(tbl.tr_lst)
    for tr in template_rows:
        tbl.remove(tr)

    queue_ = iter(stickers)
    thin_codes = {}  # dict keeps them in order without duplicates
    for sheet in range(sheets):
        for r, tr in enumerate(template_rows):
            new_tr = copy.deepcopy(tr)
            tbl.append(new_tr)
            for tc, (kind, usable_mm) in zip(new_tr.tc_lst, cells[r]):
                if kind == SLOT:
                    sticker = next(queue_, None)
                    bar_mm = _fill_slot(tc, table, sticker, usable_mm, barcode_mm, narrower)
                    if bar_mm is not None and bar_mm < MIN_BAR_MM:
                        thin_codes[sticker[1]] = None
                elif kind == OFF_PAGE:
                    _clear_cell(tc)
            if sheet and r == 0:  # every sheet starts on a new page (set after filling, which rebuilds cells)
                Paragraph(new_tr.tc_lst[0].p_lst[0], None).paragraph_format.page_break_before = True

    # An empty paragraph after the table would spill onto an extra blank page; make it tiny.
    body = doc.element.body
    for p in body.iterchildren(qn("w:p")):
        if body.index(p) > body.index(tbl) and not _p_text(p):
            fmt = Paragraph(p, None).paragraph_format
            fmt.space_before = fmt.space_after = Pt(0)
            fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            fmt.line_spacing = Pt(1)

    doc.save(dst)
    return sheets, list(thin_codes)


# ------------------------------------------------------------ Sales Order --

_SO_FORM = os.path.join("SO", "Form SO (Sales Order).xlsx")
# a form next to the program wins (so it can be changed); otherwise the copy packed inside the .exe
SO_TEMPLATE = next((p for p in (os.path.join(APP_DIR, _SO_FORM), os.path.join(BUNDLE_DIR, _SO_FORM))
                    if os.path.isfile(p)), os.path.join(APP_DIR, _SO_FORM))
SO_DRAFT_FILE = os.path.join(APP_DIR, "so_draft.json")  # the SO being scanned, kept if the app is closed
SO_FIRST_ROW, SO_ROWS = 8, 49     # item rows 8..56 on the form
SO_PRINT_AREA = "A1:K56"
SO_COL = {"tgl_dos": 3, "no": 4, "kode": 5, "qty": 6, "sat": 7, "harga": 8, "ket": 10}   # C, D, E, F, G, H, J
SO_SHRINK = ("tgl_dos", "kode", "sat", "ket")  # long text in these boxes shrinks to fit instead of spilling over
SO_LAST_COL = 11                    # the form is columns A..K; the unused copy to the right is removed
SO_STATUS_COL = 12                  # hidden column L, outside the print area: "✓" = row is on the paper
SO_STATUS_MARK = "status dicetak"   # in L1; files without it are older ones where every row was printed
# A form made since 2026-10 keeps its first item row (row 8) for kode "OP", Keterangan "PALET" (red, bold,
# Qty empty, filled in by hand); the scanned items start on row 9. L2 says so. Older SOs have no OP row and keep their layout.
SO_OP_KODE = "OP"
SO_OP_KET = "PALET"                 # Keterangan of the OP row
SO_OP_MARK = "baris 8 = OP"         # in L2 of every sheet that has the OP row
SO_CUST_MERGE = "A4:E4"             # "Cust: ..." may use A..E (up to "Jumlah SO:" in F4) and shrinks to fit
SO_HEADER = {  # cell: (label, header field)
    "A3": ("Tgl SO:", "tgl"),
    "J3": ("No. SO:", "no_so"),
    "A4": ("Cust:", "cust"),
}

_FULL_ASCII_BACK = {pair: ch for ch, pair in FULL_ASCII.items()}


def decode_full_ascii(text):
    """Turn Code 39 Full ASCII pairs back into characters, e.g. 'UT110-4 /HLSN/I' -> 'UT110-4 (LSN)'.
    Needed when the scanner's Full ASCII setting is off."""
    out, i = [], 0
    while i < len(text):
        pair = text[i:i + 2]
        if pair in _FULL_ASCII_BACK:
            out.append(_FULL_ASCII_BACK[pair])
            i += 2
        else:
            out.append(text[i])
            i += 1
    return "".join(out)


def so_page_rows(op_row):
    """How many scanned items fit on one SO paper: 49, or 48 when row 8 is the OP row."""
    return SO_ROWS - 1 if op_row else SO_ROWS


def sheet_has_op(ws):
    return ws.cell(2, SO_STATUS_COL).value == SO_OP_MARK


def write_so(template, dst, header, lines, op_row=True):
    """Fill the SO form with header {tgl, no_so, cust, salesman} and lines [{kode, qty, harga, ket}].

    op_row: row 8 of every sheet is the OP row and the lines start on row 9 (new SOs). False keeps the
    old layout (lines from row 8) for SOs made before the OP row existed.
    Every 49 (48 with the OP row) lines get their own copy of the form (a new sheet); numbering carries on.
    With the OP row, OP is No 1 and the lines are numbered from 2.
    Returns the number of pages.
    """
    import openpyxl
    from openpyxl.styles import Alignment, Font

    wb = openpyxl.load_workbook(template)
    base = wb.worksheets[0]
    for extra in wb.worksheets[1:]:  # other sheets in the form file would print as extra pages
        wb.remove(extra)
    # remove the second, unused copy of the form to the right of column K
    for merged in list(base.merged_cells.ranges):
        if merged.min_col > SO_LAST_COL:
            base.unmerge_cells(str(merged))
    if base.max_column > SO_LAST_COL:
        base.delete_cols(SO_LAST_COL + 1, base.max_column - SO_LAST_COL)
    for letter in [k for k in base.column_dimensions if openpyxl.utils.column_index_from_string(k) > SO_LAST_COL]:
        del base.column_dimensions[letter]
    # the customer's name: bold 11, allowed to use A4..E4, shrinks to fit instead of running into F4
    if not any(str(m) == SO_CUST_MERGE for m in base.merged_cells.ranges):
        base.merge_cells(SO_CUST_MERGE)
    cust_font = copy.copy(base["A4"].font)
    cust_font = Font(name=cust_font.name, size=11, bold=True, color=cust_font.color)
    base["A4"].font = cust_font
    base["A4"].alignment = Alignment(horizontal="left", vertical=base["A4"].alignment.vertical,
                                     shrink_to_fit=True)
    per_page = so_page_rows(op_row)
    pages = max(1, math.ceil(len(lines) / per_page))
    sheets = [base]
    for n in range(2, pages + 1):
        ws = wb.copy_worksheet(base)
        ws.title = f"{base.title[:24]} hal {n}"
        sheets.append(ws)

    # the template's first item row is a red example; use the plain black style of row 9 for every row
    fonts = {col: copy.copy(base.cell(SO_FIRST_ROW + 1, col).font) for col in SO_COL.values()}
    for page, ws in enumerate(sheets):
        ws.print_area = SO_PRINT_AREA
        ws.page_setup.paperSize = base.page_setup.paperSize
        ws.page_setup.orientation = base.page_setup.orientation
        ws.page_margins = copy.copy(base.page_margins)
        ws.sheet_properties.pageSetUpPr.fitToPage = True  # each form always prints on exactly 1 page
        ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1

        for cell, (label, key) in SO_HEADER.items():
            value = header.get(key, "").strip()
            ws[cell] = f"{label} {value}" if value else label
        if header.get("salesman", "").strip():
            ws["E5"] = f"Salesman : {header['salesman'].strip()}"
        if pages > 1:
            ws["F3"] = f"{ws['F3'].value}    (hal {page + 1}/{pages})"

        for r in range(SO_ROWS):
            row = SO_FIRST_ROW + r
            is_op = op_row and r == 0
            i = page * per_page + r - (1 if op_row else 0)
            line = lines[i] if not is_op and i < len(lines) else {}
            if is_op:   # only the kode and Keterangan; Qty and the rest are filled in by hand
                values = {key: None for key in SO_COL}
                values["kode"], values["ket"] = SO_OP_KODE, SO_OP_KET
            else:
                values = {"tgl_dos": line.get("tgl_dos"), "kode": line.get("kode"),
                          "qty": line.get("qty"), "sat": line.get("sat"), "harga": line.get("harga"),
                          "ket": line.get("ket")}
            # No = the row's place on the paper(s): with the OP row, OP is No 1 and the items 2, 3, ...;
            # older SOs without OP number their items from 1
            values["no"] = page * SO_ROWS + r + 1
            for key, col in SO_COL.items():
                cell = ws.cell(row, col)
                cell.value = values[key]
                cell.font = copy.copy(fonts[col])
                if is_op and key in ("kode", "ket"):
                    cell.font = Font(name=fonts[col].name, size=fonts[col].size, bold=True, color="FFFF0000")
            ws.cell(row, SO_COL["harga"]).number_format = "#,##0"
            for key in SO_SHRINK:
                cell = ws.cell(row, SO_COL[key])
                cell.alignment = Alignment(horizontal=cell.alignment.horizontal,
                                           vertical=cell.alignment.vertical, shrink_to_fit=True)
            ws.cell(row, SO_STATUS_COL).value = "✓" if line.get("printed") else None
        ws.cell(1, SO_STATUS_COL).value = SO_STATUS_MARK
        ws.cell(2, SO_STATUS_COL).value = SO_OP_MARK if op_row else None
        ws.column_dimensions[openpyxl.utils.get_column_letter(SO_STATUS_COL)].hidden = True

    wb.save(dst)
    return pages


def write_so_new_rows_only(full_so, dst, printed):
    """Make a printout of only the rows not yet on paper, from a full SO file made by write_so.

    `printed` has one True/False per SO line. A page that already went through the printer gets
    only the new rows' values (no borders, labels or old rows), so it can be fed into the printer
    again; a page with no printed rows yet gets the whole form; a page with nothing new is left out.
    Returns the number of pages to print.
    """
    import openpyxl
    from openpyxl.styles import Border, PatternFill

    if all(printed):
        raise ValueError("Tidak ada baris baru untuk dicetak.")
    wb = openpyxl.load_workbook(full_so)
    value_cols = {col for key, col in SO_COL.items() if key != "no"}  # row numbers are already on the paper
    start = 0
    for ws in list(wb.worksheets):
        op = sheet_has_op(ws)   # the OP row prints only with the whole form, never again on a reprint
        per_page = so_page_rows(op)
        flags, start = printed[start:start + per_page], start + per_page
        if all(flags):
            wb.remove(ws)  # nothing new on this page
            continue
        if not any(flags):
            continue  # this page was never printed: it needs the whole form
        new_rows = {SO_FIRST_ROW + (1 if op else 0) + r for r, done in enumerate(flags) if not done}
        for merged in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged))
        for row in ws.iter_rows(min_row=1, max_row=SO_FIRST_ROW + SO_ROWS - 1, max_col=SO_LAST_COL):
            for cell in row:
                cell.border = Border()
                cell.fill = PatternFill()
                if not (cell.row in new_rows and cell.column in value_cols):
                    cell.value = None
    wb.save(dst)
    return len(wb.worksheets)


def _after_label(value, label):
    """'Cust: Toko Maju' -> 'Toko Maju' (for reading the header of a saved SO)."""
    text = str(value or "").strip()
    return text[len(label):].strip() if text.startswith(label) else ""


def read_so(path):
    """Read an SO file made by write_so back in. Returns (header, lines) in the same shape;
    header["op"] says if the SO has the OP row (the OP row itself is not one of the lines)."""
    import datetime
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    first = wb.worksheets[0]
    has_status = first.cell(1, SO_STATUS_COL).value == SO_STATUS_MARK
    header = {key: _after_label(first[cell].value, label) for cell, (label, key) in SO_HEADER.items()}
    salesman = _after_label(first["E5"].value, "Salesman :")
    header["salesman"] = "" if salesman.count("/") >= 3 else salesman  # the form's own "Sales A/Sales B/..." list
    header["op"] = sheet_has_op(first)

    lines = []
    for ws in wb.worksheets:
        first_row = SO_FIRST_ROW + (1 if sheet_has_op(ws) else 0)   # skip the OP row
        for row in range(first_row, SO_FIRST_ROW + SO_ROWS):
            get = lambda key: ws.cell(row, SO_COL[key]).value
            kode = str(get("kode") or "").strip()
            if not kode:
                continue
            tgl = get("tgl_dos")
            if isinstance(tgl, (datetime.date, datetime.datetime)):  # Excel may turn "06/10" into a date
                tgl = tgl.strftime("%d/%m")

            def number(value):
                digits = "".join(ch for ch in str(value or "") if ch.isdigit() or ch == ".")
                try:
                    return int(float(digits)) if digits else 0
                except ValueError:
                    return 0
            lines.append({"kode": kode, "qty": number(get("qty")), "sat": str(get("sat") or "").strip(),
                          "harga": number(get("harga")), "ket": str(get("ket") or "").strip(),
                          "tgl_dos": str(tgl or "").strip(),
                          # older files have no status column: everything in them was printed
                          "printed": bool(ws.cell(row, SO_STATUS_COL).value) if has_status else True})
    return header, lines


MONTHS_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September",
             "Oktober", "November", "Desember"]


def safe_name(text):
    """Text usable as a Windows file or folder name (characters Windows doesn't allow removed)."""
    return "".join(ch for ch in text.strip() if ch not in '\\/:*?"<>|').strip(" .")


def month_year(tgl=""):
    """'Oktober_2026' for Tgl SO 'dd/mm/yyyy' - or for today if tgl is empty or can't be read."""
    try:
        _, month, year = (int(part) for part in tgl.strip().split("/"))
        return f"{MONTHS_ID[month - 1]}_{year}"
    except (ValueError, IndexError):
        return f"{MONTHS_ID[int(time.strftime('%m')) - 1]}_{time.strftime('%Y')}"


LUNAS = "_lunas"   # added to an SO's file name when the buyer has paid


def is_lunas(path):
    return os.path.splitext(os.path.basename(path))[0].lower().endswith(LUNAS)


def lunas_path(path, paid):
    """The same SO file name with '_lunas' (paid) or without it."""
    base, ext = os.path.splitext(path)
    if base.lower().endswith(LUNAS):
        base = base[:-len(LUNAS)]
    return base + (LUNAS if paid else "") + ext


def so_numbers(root, month=None):
    """SO numbers already used in one month (e.g. 'Oktober_2026'; default this month), read from the SO
    file names in every customer folder: root\\<Cust>\\<Cust>_SO_<No.>_<Month>_<Year>[_lunas].xlsx.
    The numbers start again at 1 every month; paid (_lunas) SOs keep their number. Returns
    {number: [customer folder names]}."""
    import re
    month = month or month_year()
    used = {}
    try:
        customers = [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]
    except OSError:
        return used
    for cust in customers:
        try:
            names = os.listdir(os.path.join(root, cust))   # only the files directly in the folder
        except OSError:
            continue
        for name in names:
            found = re.findall(r"_SO_(\d+)_", name)
            plain = lunas_path(name, False).lower()
            if found and plain.endswith(f"_{month}.xlsx".lower()) and \
                    os.path.isfile(os.path.join(root, cust, name)):
                used.setdefault(int(found[-1]), []).append(cust)
    return used


def next_so_number(root, month=None):
    """Next free SO number in that month (default this month): 1 for the first SO of a month."""
    return max(so_numbers(root, month), default=0) + 1


def name_part(text):
    """Text for SO file and folder names: no spaces or '-', only '_' (e.g. 'Toko-Maju Jaya' ->
    'Toko_Maju_Jaya')."""
    text = safe_name(text)
    for ch in " -\t":
        text = text.replace(ch, "_")
    while "__" in text:
        text = text.replace("__", "_")
    return text.strip("_")


def so_path(root, cust, no_so, tgl, lunas=False):
    """Where an SO is saved: root\\<Cust>\\<Cust>_SO_<No.>_<Month>_<Year>.xlsx, with '_lunas' before
    .xlsx once it is paid. Month and year come from Tgl SO (dd/mm/yyyy), or today if it can't be read."""
    cust = name_part(cust)
    return os.path.join(root, cust, f"{cust}_SO_{no_so}_{month_year(tgl)}{LUNAS if lunas else ''}.xlsx")


def so_file_name(cust):
    """File name for a customer's SO: the Cust name with characters Windows doesn't allow removed."""
    name = "".join(ch for ch in cust.strip() if ch not in '\\/:*?"<>|').strip(" .")
    return f"{name}.xlsx" if name else ""


# ------------------------------------------------------------- Settings --

SETTINGS_FILE = os.path.join(APP_DIR, "barcode_settings.json")


def read_settings():
    """Everything remembered between runs (sticker layout, Master Barang file, output folder)."""
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            saved = json.load(f)
        return saved if isinstance(saved, dict) else {}
    except (OSError, ValueError):
        return {}


def write_settings(**changes):
    saved = read_settings()
    saved.update(changes)
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(saved, f, indent=2, ensure_ascii=False)
    except OSError:
        pass  # not being able to remember a choice is not worth bothering the user about


try:  # paper sizes saved from the "Ukuran Kertas" window
    if read_settings().get("paper"):
        apply_paper(read_settings()["paper"])
except ValueError:
    pass  # a broken saved value: keep the defaults


def load_layout_setting():
    """Return (name, layout) of the sticker layout chosen last time, or the default."""
    try:
        saved = read_settings()
        name = saved["layout_name"]
        name = OLD_LAYOUT_NAMES.get(name, name)  # layouts saved by earlier versions
        name = name.replace("Custom: ", "Kustom: ").replace(" stickers, ", " stiker, ")
        layout = LAYOUTS.get(name) or saved["layout"]
        check_layout(layout)
        return name, layout
    except Exception:
        return DEFAULT_LAYOUT, LAYOUTS[DEFAULT_LAYOUT]


def save_layout_setting(name, layout):
    write_settings(layout_name=name, layout=layout)


class LayoutDialog(tk.Toplevel):
    """Small window to type in a custom sticker sheet: page size, stickers across/down, sticker size."""

    FIELDS = [  # (key, label, is_whole_number)
        ("page_w", "Lebar halaman (mm)", False),
        ("page_h", "Tinggi halaman (mm)", False),
        ("cols", "Jumlah stiker ke samping", True),
        ("rows", "Jumlah stiker ke bawah", True),
        ("sticker_w", "Lebar stiker (mm)", False),
        ("sticker_h", "Tinggi stiker (mm)", False),
        ("margin_left", "Margin kiri (mm)", False),
        ("margin_top", "Margin atas (mm)", False),
    ]

    def __init__(self, parent, layout, on_ok):
        super().__init__(parent)
        self.title("Tata letak stiker kustom")
        self.resizable(False, False)
        self.transient(parent)
        self.on_ok = on_ok
        start = {
            "page_w": layout["page_w"], "page_h": layout["page_h"],
            "cols": layout["cols"], "rows": layout["rows"],
            "sticker_w": round(min(sticker_widths(layout)), 2), "sticker_h": round(layout["sticker_h"], 2),
            "margin_left": round(layout["margin_left"] + layout["gap_left"], 2),
            "margin_top": round(layout["margin_top"] + layout["gap_top"], 2),
        }
        frame = ttk.Frame(self, padding=14)
        frame.pack(fill="both", expand=True)
        self.vars = {}
        for r, (key, label, _) in enumerate(self.FIELDS):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=f"{start[key]:g}")
            var.trace_add("write", lambda *_: self.preview())
            ttk.Entry(frame, textvariable=var, width=10, justify="right").grid(row=r, column=1, padx=8)
            self.vars[key] = var
        self.info = tk.StringVar()
        ttk.Label(frame, textvariable=self.info, foreground="#555", wraplength=260).grid(
            row=len(self.FIELDS), column=0, columnspan=2, sticky="w", pady=(10, 6))
        buttons = ttk.Frame(frame)
        buttons.grid(row=len(self.FIELDS) + 1, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Batal", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="OK", command=self.ok).pack(side="right", padx=6)
        self.preview()
        self.grab_set()

    def read(self):
        layout = {"gap_left": 0, "gap_top": 0}
        for key, label, whole in self.FIELDS:
            text = self.vars[key].get().strip().replace(",", ".")
            try:
                layout[key] = int(text) if whole else float(text)
            except ValueError:
                raise ValueError(f"'{label}' harus berupa angka{' bulat' if whole else ''}.")
        check_layout(layout)
        return layout

    def preview(self):
        try:
            layout = self.read()
        except ValueError as e:
            self.info.set(str(e))
            return
        self.info.set(f"{stickers_per_sheet(layout)} stiker per lembar. Lebar untuk barcode "
                      f"{barcode_space_mm(layout):.1f} mm.")

    def ok(self):
        try:
            layout = self.read()
        except ValueError as e:
            messagebox.showerror("Tata letak stiker kustom", str(e), parent=self)
            return
        name = (f"Kustom: {layout['cols']} x {layout['rows']} stiker, "
                f"{layout['sticker_w']:g} x {layout['sticker_h']:g} mm")
        self.destroy()
        self.on_ok(name, layout)


class PaperDialog(tk.Toplevel):
    """'Ukuran Kertas' window (gear button): the sticker paper drawn to scale, every box showing its size.
    Click boxes to select them (Ctrl+click adds/removes, Shift+click selects a rectangle), then type
    Lebar / Tinggi below and press Terapkan. On the paper all boxes of a column are equally wide and all
    boxes of a row equally high, so Lebar changes the selected boxes' columns and Tinggi their rows."""

    MAX_W, MAX_H = 900, 540   # room for the drawing (pixels)
    PAD_L, PAD_T = 50, 24
    SELECTED = "#9cc3f0"

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Ukuran Kertas Stiker")
        self.resizable(False, False)
        self.transient(app)
        self.redraw_job = None
        self.cols = list(PAPER["cols"])
        self.rows = list(PAPER["rows"])
        self.selected = set()      # (row, col) of the selected boxes
        self.anchor = None         # box clicked last, for Shift+click
        self.box_at = {}           # (row, col) -> (x1, y1, x2, y2) on the canvas

        top = ttk.Frame(self, padding=(12, 10, 12, 0))
        top.pack(fill="x")
        self.size_var = tk.StringVar()
        ttk.Label(top, textvariable=self.size_var, style="Summary.TLabel").pack(side="left")
        ttk.Label(top, text="Klik kotak untuk memilih (Ctrl+klik = tambah, Shift+klik = beberapa sekaligus).",
                  foreground="#666").pack(side="right")

        self.canvas = tk.Canvas(self, width=self.MAX_W + self.PAD_L + 70, height=self.MAX_H + self.PAD_T + 50,
                                bg="#f4f4f4", highlightthickness=0)
        self.canvas.pack(padx=12, pady=8)
        self.canvas.bind("<Button-1>", self.on_click)

        def var(value):
            v = tk.StringVar(value=f"{value:g}")
            v.trace_add("write", lambda *_: self.schedule_redraw())
            return v
        self.right_var, self.bottom_var = var(PAPER["right_empty"]), var(PAPER["bottom_empty"])
        self.left_var, self.top_var = var(PAPER["left"]), var(PAPER["top"])

        def entry(v):
            return self.canvas.create_window(0, 0, window=ttk.Entry(self.canvas, textvariable=v, width=5,
                                                                    justify="center"))
        self.right_win, self.bottom_win = entry(self.right_var), entry(self.bottom_var)

        edit = ttk.LabelFrame(self, text="Ukuran kotak yang dipilih", padding=8)
        edit.pack(fill="x", padx=12)
        self.sel_var = tk.StringVar(value="Belum ada kotak dipilih.")
        ttk.Label(edit, textvariable=self.sel_var, width=34).pack(side="left")
        ttk.Label(edit, text="Lebar:").pack(side="left", padx=(10, 2))
        self.lebar_var = tk.StringVar()
        lebar = ttk.Entry(edit, textvariable=self.lebar_var, width=7, justify="center")
        lebar.pack(side="left")
        ttk.Label(edit, text="Tinggi:").pack(side="left", padx=(14, 2))
        self.tinggi_var = tk.StringVar()
        tinggi = ttk.Entry(edit, textvariable=self.tinggi_var, width=7, justify="center")
        tinggi.pack(side="left")
        ttk.Label(edit, text="mm").pack(side="left", padx=(4, 12))
        ttk.Button(edit, text="Terapkan", command=self.apply_size).pack(side="left")
        ttk.Button(edit, text="Pilih semua", command=self.select_all).pack(side="left", padx=6)
        for e in (lebar, tinggi):
            e.bind("<Return>", lambda ev: self.apply_size())

        shift = ttk.Frame(self, padding=(12, 8, 12, 0))
        shift.pack(fill="x")
        ttk.Label(shift, text="Geser semua kotak ke kanan:").pack(side="left")
        ttk.Entry(shift, textvariable=self.left_var, width=6, justify="center").pack(side="left", padx=(4, 14))
        ttk.Label(shift, text="ke bawah:").pack(side="left")
        ttk.Entry(shift, textvariable=self.top_var, width=6, justify="center").pack(side="left", padx=4)
        ttk.Label(shift, text="(biasanya 0)", foreground="#888").pack(side="left", padx=6)

        buttons = ttk.Frame(self, padding=12)
        buttons.pack(fill="x")
        self.msg = tk.Label(buttons, text="", fg="#c62828", anchor="w")
        self.msg.pack(side="left", fill="x", expand=True)
        ttk.Button(buttons, text="Batal", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Simpan", style="Big.TButton", command=self.save).pack(side="right", padx=6)
        ttk.Button(buttons, text="Ukuran awal", command=self.reset).pack(side="right", padx=6)

        self.redraw()
        self.grab_set()

    # ---- selecting boxes

    def on_click(self, event):
        box = next((rc for rc, (x1, y1, x2, y2) in self.box_at.items()
                    if x1 <= event.x <= x2 and y1 <= event.y <= y2), None)
        if box is None:
            return
        if event.state & SHIFT_MASK and self.anchor:   # every box in the rectangle between anchor and here
            (r1, c1), (r2, c2) = self.anchor, box
            self.selected |= {(r, c) for r in range(min(r1, r2), max(r1, r2) + 1)
                              for c in range(min(c1, c2), max(c1, c2) + 1)}
        elif event.state & CTRL_MASK:
            self.selected ^= {box}
            self.anchor = box
        else:
            self.selected = {box}
            self.anchor = box
        self.show_selection()
        self.redraw()

    def select_all(self):
        self.selected = {(r, c) for r in range(len(self.rows)) for c in range(len(self.cols))}
        self.show_selection()
        self.redraw()

    def show_selection(self):
        """Put the selected boxes' size in the Lebar/Tinggi fields (empty if they differ)."""
        n = len(self.selected)
        self.sel_var.set("Belum ada kotak dipilih." if not n else f"{n} kotak dipilih.")
        widths = {self.cols[c] for _, c in self.selected}
        heights = {self.rows[r] for r, _ in self.selected}
        self.lebar_var.set(f"{widths.pop():g}" if len(widths) == 1 else "")
        self.tinggi_var.set(f"{heights.pop():g}" if len(heights) == 1 else "")

    def apply_size(self):
        if not self.selected:
            self.msg.config(text="Pilih kotak dulu (klik kotaknya).")
            return
        values = {}
        for name, v in (("Lebar", self.lebar_var), ("Tinggi", self.tinggi_var)):
            text = v.get().strip().replace(",", ".")
            if text:   # an empty field leaves that size as it is
                try:
                    values[name] = float(text)
                except ValueError:
                    self.msg.config(text=f"{name} '{v.get()}' bukan angka.")
                    return
        cols, rows = list(self.cols), list(self.rows)
        for r, c in self.selected:
            if "Lebar" in values:
                cols[c] = values["Lebar"]
            if "Tinggi" in values:
                rows[r] = values["Tinggi"]
        try:
            check_paper(dict(self.read(), cols=cols, rows=rows))
        except ValueError as e:
            self.msg.config(text=str(e))
            return
        self.cols, self.rows = cols, rows
        self.msg.config(text="")
        self.redraw()

    # ---- drawing

    def schedule_redraw(self):
        if self.redraw_job:
            self.after_cancel(self.redraw_job)
        self.redraw_job = self.after(250, self.redraw)

    def read(self):
        def num(v):
            try:
                return float(v.get().strip().replace(",", "."))
            except ValueError:
                raise ValueError(f"'{v.get()}' bukan angka - semua ukuran harus angka (mm).")
        return check_paper({"cols": self.cols, "rows": self.rows,
                            "right_empty": num(self.right_var), "bottom_empty": num(self.bottom_var),
                            "left": num(self.left_var), "top": num(self.top_var)})

    def redraw(self):
        self.redraw_job = None
        try:
            paper = self.read()
        except ValueError as e:
            self.msg.config(text=str(e))
            return
        c = self.canvas
        c.delete("drawing")
        page_w = paper["left"] + sum(paper["cols"]) + paper["right_empty"]
        page_h = paper["top"] + sum(paper["rows"]) + paper["bottom_empty"]
        portrait = LAYOUTS[_ASLI].get("portrait")
        self.size_var.set(f"Kertas: {page_w:g} x {page_h:g} mm" +
                          (f"   (dicetak potret: {page_h:g} x {page_w:g} mm)" if portrait else ""))
        scale = min(self.MAX_W / page_w, self.MAX_H / page_h)
        ox, oy = self.PAD_L, self.PAD_T
        X = lambda mm: ox + mm * scale
        Y = lambda mm: oy + mm * scale

        c.create_rectangle(X(0), Y(0), X(page_w), Y(page_h), fill="#d9d9d9", outline="#555",
                           width=2, tags="drawing")   # the paper; grey = empty paper
        if portrait:
            c.create_text(X(0) - 25, Y(page_h / 2), text="sisi ini masuk printer duluan  ->", angle=90,
                          fill="#888", font=("Segoe UI", 9), tags="drawing")
        self.box_at = {}
        y = paper["top"]
        for r, h in enumerate(paper["rows"]):
            x = paper["left"]
            for col, w in enumerate(paper["cols"]):
                box = (X(x), Y(y), X(x + w), Y(y + h))
                self.box_at[r, col] = box
                chosen = (r, col) in self.selected
                c.create_rectangle(*box, fill=self.SELECTED if chosen else "white",
                                   outline="#1f5fa8" if chosen else "#888", width=2 if chosen else 1,
                                   tags="drawing")
                c.create_text((box[0] + box[2]) / 2, (box[1] + box[3]) / 2,
                              text=f"Lebar: {w:g}\nTinggi: {h:g}", justify="center",
                              fill="#222", font=("Segoe UI", 9), tags="drawing")
                x += w
            y += h

        right_x0 = paper["left"] + sum(paper["cols"])
        if paper["right_empty"] > 0:   # empty paper on the right
            c.create_text((X(right_x0) + X(page_w)) / 2, Y(page_h / 2) + 30, text="kosong", angle=90,
                          fill="#666", font=("Segoe UI", 9), tags="drawing")
        c.coords(self.right_win, (X(right_x0) + X(page_w)) / 2 if paper["right_empty"] * scale > 50
                 else X(page_w) + 30, Y(page_h / 2) - 22)
        c.create_text(X(page_w) + 30, Y(page_h / 2) - 44, text="kosong kanan:", fill="#666",
                      font=("Segoe UI", 9), tags="drawing")
        bottom_y0 = paper["top"] + sum(paper["rows"])
        c.coords(self.bottom_win, (X(0) + X(page_w)) / 2, Y(page_h) + 22)  # empty strip at the bottom
        c.create_text((X(0) + X(page_w)) / 2 - 30, Y(page_h) + 22, text="kosong bawah:", anchor="e",
                      fill="#666", font=("Segoe UI", 9), tags="drawing")
        if paper["bottom_empty"] > 0:
            c.create_line(X(page_w / 2), Y(bottom_y0), X(page_w / 2), Y(page_h) + 12, fill="#999",
                          arrow="first", tags="drawing")
        c.tag_raise(self.right_win)
        c.tag_raise(self.bottom_win)

    def reset(self):
        self.cols, self.rows = list(DEFAULT_PAPER["cols"]), list(DEFAULT_PAPER["rows"])
        for v, key in ((self.right_var, "right_empty"), (self.bottom_var, "bottom_empty"),
                       (self.left_var, "left"), (self.top_var, "top")):
            v.set(f"{DEFAULT_PAPER[key]:g}")
        self.show_selection()
        self.redraw()

    def save(self):
        try:
            paper = self.read()
        except ValueError as e:
            messagebox.showerror("Ukuran Kertas", str(e), parent=self)
            return
        apply_paper(paper)
        write_settings(paper=paper)
        self.app.show_layout()
        self.app.update_summary()
        self.app.status_var.set("Ukuran kertas disimpan - berlaku untuk cetak stiker berikutnya.")
        self.destroy()


# ------------------------------------------------------------ Dashboard --

COLUMNS = [  # (id, heading, width, anchor)
    ("check", "✓", 36, "center"),
    ("kode", "Kode User", 130, "w"),
    ("nama", "Nama Barang", 340, "w"),
    ("harga", "Harga", 90, "e"),
    ("sat", "Sat", 50, "center"),
    ("kodeharga", "Kode Harga", 110, "w"),
    ("print", "Cetak", 60, "center"),
]
CHECK_COL = "#1"
PRINT_COL = f"#{len(COLUMNS)}"
SHIFT_MASK, CTRL_MASK = 0x0001, 0x0004  # keyboard modifier bits in a Tk event's state


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pembuat Barcode")
        self.geometry("1050x720")
        self.minsize(860, 580)
        self.events = queue.Queue()
        self.items = []          # (kode, nama, harga, sat) from Master Barang
        self.by_kode = {}        # Kode User (upper case) -> item index, for the scanner
        self.qty = {}            # item index -> how many stickers (ticked items only)
        self.output_dir = None
        self.editor = None       # the Print box being edited, if any
        self.search_job = None
        self.sort_col, self.sort_reverse = None, False  # list order; None = Master Barang order
        self.units_cache = {}    # kode -> barcode width in narrow-bar units
        self.layout_name, self.layout = load_layout_setting()

        style = ttk.Style(self)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=8)
        style.configure("Gear.TButton", font=("Segoe UI Symbol", 16), padding=(4, 3))
        style.configure("Summary.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", rowheight=24)

        style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=(18, 4))
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=4, pady=4)
        root = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(root, text="Cetak Stiker")
        self.so_tab = SOTab(self.tabs, self)
        self.tabs.add(self.so_tab, text="SO")
        self.tabs.bind("<<NotebookTabChanged>>", lambda e: self.so_tab.on_shown()
                       if self.tabs.select() == str(self.so_tab) else None)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

        ttk.Label(root, text="Pembuat Barcode", style="Title.TLabel").grid(row=0, column=0, sticky="w")

        # ---- files
        files = ttk.Frame(root)
        files.grid(row=1, column=0, sticky="ew", pady=(6, 8))
        files.columnconfigure(1, weight=1)
        self.master_var = tk.StringVar()
        self.out_var = tk.StringVar()
        for r, (label, var, cmd) in enumerate((
                ("File Master Barang:", self.master_var, self.pick_master),
                ("Simpan hasil ke:", self.out_var, self.pick_output))):
            ttk.Label(files, text=label).grid(row=r, column=0, sticky="w", pady=2)
            ttk.Entry(files, textvariable=var).grid(row=r, column=1, sticky="ew", padx=6)
            ttk.Button(files, text="Pilih...", command=cmd).grid(row=r, column=2)
        ttk.Label(files, text="Tata letak stiker:").grid(row=2, column=0, sticky="w", pady=2)
        layout_row = ttk.Frame(files)
        layout_row.grid(row=2, column=1, columnspan=2, sticky="ew", padx=6)
        self.layout_var = tk.StringVar(value=self.layout_name)
        self.layout_box = ttk.Combobox(layout_row, textvariable=self.layout_var, state="readonly", width=42)
        self.layout_box.pack(side="left")
        self.layout_box.bind("<<ComboboxSelected>>", self.on_layout_selected)
        self.layout_info = tk.StringVar()
        ttk.Label(layout_row, textvariable=self.layout_info, foreground="#666").pack(side="left", padx=10)
        self.show_layout()

        # ---- search bar
        bar = ttk.Frame(root)
        bar.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(bar, text="Cari:").pack(side="left")
        self.search_var = tk.StringVar()
        search = ttk.Entry(bar, textvariable=self.search_var, width=40)
        search.pack(side="left", padx=6)
        self.search_var.trace_add("write", lambda *_: self.schedule_refresh())
        self.only_ticked = tk.BooleanVar()
        ttk.Checkbutton(bar, text="Tampilkan yang dicentang saja", variable=self.only_ticked,
                        command=self.refresh_tree).pack(side="left", padx=10)
        ttk.Button(bar, text="Hapus semua centang", command=self.untick_all).pack(side="left")
        ttk.Label(bar, text="Cetak untuk baris terpilih:").pack(side="left", padx=(20, 4))
        self.bulk_var = tk.StringVar(value="1")
        bulk = ttk.Entry(bar, textvariable=self.bulk_var, width=5, justify="center")
        bulk.pack(side="left")
        bulk.bind("<Return>", lambda e: self.apply_bulk_qty())
        ttk.Button(bar, text="Terapkan", command=self.apply_bulk_qty).pack(side="left", padx=4)
        self.count_var = tk.StringVar()

        # ---- item list
        table = ttk.Frame(root)
        table.grid(row=3, column=0, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=[c[0] for c in COLUMNS], show="headings", selectmode="extended")
        for cid, heading, width, anchor in COLUMNS:
            self.tree.heading(cid, text=heading, command=lambda c=cid: self.sort_by(c))
            self.tree.column(cid, width=width, anchor=anchor, stretch=(cid == "nama"))
        # alternating row colours, ticked rows green, selected rows blue. Row colours hide Tk's own
        # selection colour, so selected rows get the "sel" colour by hand (see show_selection).
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd", background="#ececec")
        self.tree.tag_configure("ticked", background="#cdeccd")
        self.tree.tag_configure("sel", background="#9cc3f0")
        self.sel_shown = set()   # rows currently coloured as selected
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.show_selection())
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.on_scroll)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<space>", self.on_space)
        self.tree.bind("<Key>", self.on_key)
        self.tree.bind("<Control-a>", self.select_all)
        self.tree.bind("<MouseWheel>", lambda e: self.close_editor())
        search.bind("<Down>", lambda e: self.focus_list())
        hint = ttk.Frame(root)
        hint.grid(row=4, column=0, sticky="ew", pady=(4, 0))
        ttk.Label(hint, textvariable=self.count_var, foreground="#666").pack(side="right", anchor="n")
        ttk.Label(hint, text="Klik ☐ untuk mencentang, klik kolom Cetak untuk mengetik jumlah.  "
                             "Beberapa baris: Shift+klik atau Ctrl+klik (Ctrl+A = semua), lalu klik salah satu "
                             "kotak ☐ / Cetak-nya, ketik angka, atau pakai 'Cetak untuk baris terpilih'.",
                  foreground="#666", wraplength=760, justify="left").pack(side="left")

        # ---- bottom bar
        bottom = ttk.Frame(root)
        bottom.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        # the buttons are packed first so they always keep their room; the summary gets what is left
        self.open_btn = ttk.Button(bottom, text="Buka folder hasil", command=self.open_output, state="disabled")
        self.open_btn.pack(side="right")
        self.all_btn = ttk.Button(bottom, text="Ekspor semua label", command=self.start_export_all)
        self.all_btn.pack(side="right", padx=6)
        self.print_btn = ttk.Button(bottom, text="Cetak Stiker", style="Big.TButton", command=self.start_print)
        self.print_btn.pack(side="right", padx=6)
        ttk.Button(bottom, text="⚙", style="Gear.TButton", width=3,
                   command=lambda: PaperDialog(self)).pack(side="right")   # paper size window
        summary = ttk.Frame(bottom)
        summary.pack(side="left", fill="x", expand=True)
        self.summary_var = tk.StringVar(value="Belum ada yang dicentang.")
        ttk.Label(summary, textvariable=self.summary_var, style="Summary.TLabel").pack(anchor="w")
        self.warn_var = tk.StringVar()  # "codes too long" warning, on its own line
        ttk.Label(summary, textvariable=self.warn_var, foreground="#b45309").pack(anchor="w")

        status = ttk.Frame(root)
        status.grid(row=6, column=0, sticky="ew", pady=(6, 0))
        self.status_var = tk.StringVar()
        ttk.Label(status, textvariable=self.status_var).pack(side="left")
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=160)

        # the files chosen last time (e.g. a Master Barang in a shared folder), else the ones next to the app
        saved = read_settings()
        master = saved.get("master_path", "")
        if not os.path.isfile(master):
            master = os.path.join(APP_DIR, "MASTER BARANG UPDATE.xls")
        if not os.path.isfile(master):   # public demo: made-up items
            master = os.path.join(APP_DIR, "demo", "MASTER BARANG DEMO.xlsx")
        if os.path.isfile(master):
            self.master_var.set(master)
        self.out_var.set(saved.get("output_dir") or os.path.join(APP_DIR, "barcode_output"))

        self.after(100, self.poll_events)
        if self.master_var.get():
            self.load_master()
        else:
            self.status_var.set("Pilih file Master Barang untuk mulai.")

    # ---- sticker layout

    def show_layout(self):
        names = list(LAYOUTS)
        if self.layout_name not in LAYOUTS:
            names.append(self.layout_name)  # the saved custom layout
        self.layout_box["values"] = names + [CUSTOM]
        self.layout_var.set(self.layout_name)
        widths = sticker_widths(self.layout)
        page_w, page_h = self.layout["page_w"], self.layout["page_h"]
        if self.layout.get("portrait"):  # printed turned a quarter turn
            page_w, page_h = page_h, page_w
        self.layout_info.set(f"{stickers_per_sheet(self.layout)} per lembar  ·  halaman "
                             f"{page_w:g} x {page_h:g} mm" + (" (potret)" if self.layout.get("portrait") else "")
                             + f"  ·  stiker {round(min(widths), 1):g} x {self.layout['sticker_h']:.0f} mm")

    def on_layout_selected(self, event=None):
        name = self.layout_var.get()
        if name == CUSTOM:
            LayoutDialog(self, self.layout, self.set_layout)
            self.layout_var.set(self.layout_name)  # until the dialog is confirmed
        elif name in LAYOUTS:
            self.set_layout(name, LAYOUTS[name])

    def set_layout(self, name, layout):
        self.layout_name, self.layout = name, layout
        save_layout_setting(name, layout)
        self.show_layout()
        self.update_summary()

    # ---- files

    def pick_master(self):
        path = filedialog.askopenfilename(title="Pilih file Master Barang",
                                          filetypes=[("File Excel", "*.xls *.xlsx"), ("Semua file", "*.*")])
        if path:
            self.master_var.set(path)
            write_settings(master_path=path)
            self.load_master()

    def pick_output(self):
        path = filedialog.askdirectory(title="Pilih folder hasil")
        if path:
            self.out_var.set(path)
            write_settings(output_dir=path)

    def open_output(self):
        if self.output_dir and os.path.isdir(self.output_dir):
            os.startfile(self.output_dir)

    def load_master(self):
        path = self.master_var.get().strip()
        if not os.path.isfile(path):
            messagebox.showerror("Pembuat Barcode", "Silakan pilih file Master Barang.")
            return
        self.set_busy(True, "Membaca Master Barang...")

        def job():
            items, skipped = load_items(path)
            self.events.put(("loaded", items, skipped))
        self.run_in_background(job)

    # ---- item list

    def item_values(self, idx):
        kode, nama, harga, sat = self.items[idx]
        q = self.qty.get(idx, 0)
        return ("☑" if q else "☐", kode, nama, rupiah(harga), sat, kode_harga(harga, sat), q or "")

    def visible_indexes(self):
        words = self.search_var.get().lower().split()
        only_ticked = self.only_ticked.get()
        for idx, (kode, nama, _, _) in enumerate(self.items):
            if only_ticked and idx not in self.qty:
                continue
            if words:
                text = f"{kode} {nama}".lower()
                if not all(w in text for w in words):
                    continue
            yield idx

    def schedule_refresh(self):
        if self.search_job:
            self.after_cancel(self.search_job)
        self.search_job = self.after(250, self.refresh_tree)

    def refresh_tree(self):
        self.search_job = None
        self.close_editor()
        self.tree.delete(*self.tree.get_children())
        self.sel_shown = set()
        indexes = list(self.visible_indexes())
        if self.sort_col:
            indexes.sort(key=self.sort_key(self.sort_col), reverse=self.sort_reverse)
        shown = 0
        for idx in indexes:
            self.tree.insert("", "end", iid=str(idx), values=self.item_values(idx), tags=self.row_tags(idx, shown))
            shown += 1
        self.count_var.set(f"Menampilkan {angka(shown)} dari {angka(len(self.items))} barang")

    def sort_key(self, col):
        """Key function for sorting item indexes by a column. The first click on a header gives the
        most useful order: ticked items first, biggest Print first, A-Z / cheapest first for the rest."""
        items, qty = self.items, self.qty
        return {
            "check": lambda i: i not in qty,
            "kode": lambda i: items[i][0].lower(),
            "nama": lambda i: items[i][1].lower(),
            "harga": lambda i: items[i][2],
            "sat": lambda i: items[i][3].lower(),
            "kodeharga": lambda i: kode_harga(items[i][2], items[i][3]),
            "print": lambda i: -qty.get(i, 0),
        }[col]

    def sort_by(self, col):
        self.close_editor()
        if self.sort_col == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_col, self.sort_reverse = col, False
        for cid, heading, _, _ in COLUMNS:  # show an arrow on the sorted column
            arrow = (" ▼" if self.sort_reverse else " ▲") if cid == col else ""
            self.tree.heading(cid, text=heading + arrow)
        self.refresh_tree()

    def show_selection(self):
        selected = set(self.tree.selection())
        for iid in self.sel_shown - selected:
            if self.tree.exists(iid):
                self.tree.item(iid, tags=self.row_tags(int(iid), self.tree.index(iid), selected=False))
        for iid in selected - self.sel_shown:
            self.tree.item(iid, tags=("sel",))
        self.sel_shown = selected

    def row_tags(self, idx, position, selected=None):
        if selected is None:
            selected = str(idx) in self.sel_shown
        if selected:
            return ("sel",)
        if idx in self.qty:
            return ("ticked",)
        return ("odd",) if position % 2 else ("even",)

    def update_row(self, idx):
        iid = str(idx)
        if self.tree.exists(iid):
            self.tree.item(iid, values=self.item_values(idx), tags=self.row_tags(idx, self.tree.index(iid)))

    def set_qty(self, idx, q):
        self.set_many({idx: q})

    def set_many(self, changes):
        """Apply {item index: print count} (0 = untick) and refresh the screen once."""
        for idx, q in changes.items():
            if q > 0:
                self.qty[idx] = q
            else:
                self.qty.pop(idx, None)
        if len(changes) > 200:
            selection = self.tree.selection()
            self.refresh_tree()  # quicker than updating thousands of rows one by one
            self.tree.selection_set([i for i in selection if self.tree.exists(i)])
        else:
            for idx in changes:
                self.update_row(idx)
        self.update_summary()

    def toggle(self, idx):
        self.set_qty(idx, 0 if idx in self.qty else 1)

    def tick_group(self, idxs):
        """Tick all of them (keeping their Print count), unless all are ticked already: then untick all."""
        tick = any(i not in self.qty for i in idxs)
        self.set_many({i: (self.qty.get(i, 1) if tick else 0) for i in idxs})

    def click_group(self, iid):
        """Rows a click on `iid` acts on: the whole selection if `iid` is part of it, else just `iid`."""
        selected = self.tree.selection()
        if iid in selected and len(selected) > 1:
            return [int(i) for i in selected]
        return [int(iid)]

    def selected_indexes(self):
        return [int(i) for i in self.tree.selection()]

    def apply_bulk_qty(self):
        self.close_editor()
        idxs = self.selected_indexes()
        if not idxs:
            messagebox.showinfo("Pembuat Barcode", "Pilih baris dulu: klik satu baris, lalu Shift+klik atau "
                                                 "Ctrl+klik untuk menambah (Ctrl+A memilih semua yang tampil).")
            return
        text = self.bulk_var.get().strip()
        if not text.isdigit():
            messagebox.showerror("Pembuat Barcode", "Ketik jumlah stiker (angka bulat, 0 = hapus centang).")
            return
        self.set_many({i: min(int(text), 9999) for i in idxs})
        self.tree.focus_set()

    def untick_all(self):
        if self.qty and messagebox.askyesno("Pembuat Barcode", f"Hapus centang dari {len(self.qty)} barang?"):
            self.qty.clear()
            self.refresh_tree()
            self.update_summary()

    def on_click(self, event):
        self.close_editor()
        if event.state & (SHIFT_MASK | CTRL_MASK):
            return  # Shift/Ctrl+click: just select rows (Treeview's normal behaviour)
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid, col = self.tree.identify_row(event.y), self.tree.identify_column(event.x)
        if col not in (CHECK_COL, PRINT_COL):
            return
        group = self.click_group(iid)
        if col == CHECK_COL:
            self.tick_group(group)
        else:
            self.after_idle(lambda: self.edit_qty(iid, group))
        if len(group) > 1:
            return "break"  # keep the selection instead of selecting only the clicked row

    def on_double_click(self, event):
        if event.state & (SHIFT_MASK | CTRL_MASK):
            return
        iid, col = self.tree.identify_row(event.y), self.tree.identify_column(event.x)
        if iid and col not in (CHECK_COL, PRINT_COL):
            self.toggle(int(iid))

    def on_space(self, event):
        idxs = self.selected_indexes()
        if idxs:
            self.tick_group(idxs)
        return "break"

    def on_key(self, event):
        """Typing a number (or Enter) edits the Print count of the selected rows."""
        iid = self.tree.focus()
        if not iid or not self.tree.selection():
            return
        if event.char.isdigit():
            self.edit_qty(iid, self.click_group(iid), start_text=event.char)
            return "break"
        if event.keysym in ("Return", "KP_Enter"):
            self.edit_qty(iid, self.click_group(iid))
            return "break"

    def select_all(self, event=None):
        self.tree.selection_set(self.tree.get_children())
        return "break"

    def on_scroll(self, *args):
        self.close_editor()
        self.tree.yview(*args)

    def focus_list(self):
        children = self.tree.get_children()
        if children:
            self.tree.focus_set()
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])

    def edit_qty(self, iid, targets=None, start_text=None):
        """Show a Print box over row `iid`; what is typed is applied to all `targets` rows."""
        self.tree.see(iid)
        self.tree.update_idletasks()
        bbox = self.tree.bbox(iid, PRINT_COL)
        if not bbox:
            return
        x, y, w, h = bbox
        targets = targets or [int(iid)]
        var = tk.StringVar(value=start_text if start_text is not None else str(self.qty.get(int(iid), 1)))
        entry = ttk.Entry(self.tree, textvariable=var, justify="center")
        entry.place(x=x, y=y, width=w, height=h)
        if start_text is None:
            entry.select_range(0, "end")
        else:
            entry.icursor("end")
        entry.focus_set()
        self.editor = (entry, targets, var)
        if len(targets) > 1:
            self.status_var.set(f"Ketik jumlah Cetak untuk {len(targets)} baris terpilih, lalu tekan Enter.")
        entry.bind("<Return>", lambda e: self.close_editor())
        entry.bind("<FocusOut>", lambda e: self.close_editor())
        entry.bind("<Escape>", lambda e: self.close_editor(save=False))

    def close_editor(self, save=True):
        if not self.editor:
            return
        entry, targets, var = self.editor
        self.editor = None
        if save:
            text = var.get().strip()
            if text.isdigit() or text == "":
                q = min(int(text), 9999) if text else 0
                self.set_many({i: q for i in targets})
                if len(targets) > 1:
                    self.status_var.set(f"Cetak diisi {q} untuk {len(targets)} baris.")
        entry.destroy()
        self.tree.focus_set()

    def update_summary(self):
        if not self.qty:
            self.summary_var.set("Belum ada yang dicentang.")
            self.warn_var.set("")
            return
        total = sum(self.qty.values())
        per_sheet = stickers_per_sheet(self.layout)
        sheets = math.ceil(total / per_sheet)
        text = (f"{len(self.qty)} barang dicentang  ·  {total} stiker  ·  "
                f"{sheets} lembar ({per_sheet} per lembar)")
        space = barcode_space_mm(self.layout)
        too_long = sum(space / self.barcode_units(self.items[i][0]) < MIN_BAR_MM for i in self.qty)
        self.warn_var.set(f"⚠ {too_long} kode terlalu panjang untuk ukuran stiker ini (garis barcode tipis)"
                          if too_long else "")
        self.summary_var.set(text)

    def barcode_units(self, kode):
        if kode not in self.units_cache:
            self.units_cache[kode] = barcode_image(kode, narrow_px=8, height_px=1, quiet=0)[1]
        return self.units_cache[kode]

    # ---- printing / export

    def start_print(self):
        self.close_editor()
        if not self.qty:
            messagebox.showinfo("Pembuat Barcode", "Centang dulu barang yang mau dicetak.")
            return
        out = self.out_var.get().strip()
        if not out:
            messagebox.showerror("Pembuat Barcode", "Silakan pilih tempat menyimpan hasil.")
            return
        stickers = []
        for idx in sorted(self.qty):  # same order as the Master Barang list
            kode, _, harga, sat = self.items[idx]
            stickers += [(kode_harga(harga, sat), kode)] * self.qty[idx]
        self.set_busy(True, f"Membuat {len(stickers)} stiker...")
        layout = dict(self.layout)

        def job():
            os.makedirs(out, exist_ok=True)
            dst = os.path.join(out, f"stickers_{time.strftime('%Y%m%d_%H%M%S')}.docx")
            sheets, thin = make_sticker_doc(dst, stickers, layout)
            self.events.put(("printed", dst, len(stickers), sheets, thin))
        self.run_in_background(job)

    def start_export_all(self):
        if not self.items:
            messagebox.showerror("Pembuat Barcode", "Muat file Master Barang dulu.")
            return
        out = self.out_var.get().strip()
        if not out:
            messagebox.showerror("Pembuat Barcode", "Silakan pilih tempat menyimpan hasil.")
            return
        if not messagebox.askyesno("Pembuat Barcode", f"Buat label A4 untuk semua {angka(len(self.items))} barang?\n"
                                                    "Ini memakan waktu sekitar satu menit."):
            return
        items = list(self.items)
        self.set_busy(True, "Mengekspor semua label...")

        def job():
            os.makedirs(out, exist_ok=True)
            write_csv(items, os.path.join(out, "barcode_data.csv"))
            chunks = [items[i:i + LABELS_PER_FILE] for i in range(0, len(items), LABELS_PER_FILE)]
            for n, chunk in enumerate(chunks, 1):
                self.events.put(("status", f"Mengekspor semua label: file Word {n} dari {len(chunks)}..."))
                write_docx(chunk, os.path.join(out, f"barcode_labels_part{n:02d}.docx"))
            self.events.put(("exported", out, len(chunks)))
        self.run_in_background(job)

    # ---- background work (runs in a thread, talks to the window through self.events)

    def set_busy(self, busy, message=""):
        state = "disabled" if busy else "normal"
        self.print_btn.configure(state=state)
        self.all_btn.configure(state=state)
        if busy:
            self.status_var.set(message)
            self.progress.pack(side="left", padx=10)
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.pack_forget()

    def run_in_background(self, job):
        def runner():
            try:
                job()
            except PermissionError as e:
                self.events.put(("error", "Tidak bisa menyimpan file - apakah sedang terbuka di Word atau Excel? "
                                          f"Tutup dulu lalu coba lagi.\n\n{e}"))
            except Exception as e:
                self.events.put(("error", str(e)))
        threading.Thread(target=runner, daemon=True).start()

    def poll_events(self):
        try:
            while True:
                ev = self.events.get_nowait()
                kind = ev[0]
                if kind == "status":
                    self.status_var.set(ev[1])
                elif kind == "loaded":
                    self.items, self.qty = ev[1], {}
                    self.by_kode = {item[0].upper(): i for i, item in enumerate(self.items)}
                    self.set_busy(False)
                    self.status_var.set(f"{angka(len(self.items))} barang dimuat"
                                        + (f" ({ev[2]} tanpa Kode User dilewati)." if ev[2] else "."))
                    self.refresh_tree()
                    self.update_summary()
                elif kind == "printed":
                    _, path, n, sheets, thin = ev
                    self.output_dir = os.path.dirname(path)
                    self.set_busy(False)
                    self.open_btn.configure(state="normal")
                    self.status_var.set(f"{n} stiker dibuat di {sheets} lembar: {os.path.basename(path)}. "
                                        "Dibuka di Word - tekan Ctrl+P di sana untuk mencetak.")
                    os.startfile(path)
                    if thin:
                        messagebox.showwarning(
                            "Pembuat Barcode",
                            "Kode-kode ini panjang, jadi garis barcode-nya sangat tipis dan mungkin tidak terbaca scanner. "
                            "Coba scan dulu sebelum mencetak banyak:\n\n" + "\n".join(thin[:20])
                            + (f"\n... dan {len(thin) - 20} lainnya" if len(thin) > 20 else ""))
                elif kind == "exported":
                    self.output_dir = ev[1]
                    self.set_busy(False)
                    self.open_btn.configure(state="normal")
                    self.status_var.set(f"barcode_data.csv dan {ev[2]} file Word diekspor ke {ev[1]}")
                elif kind == "error":
                    self.set_busy(False)
                    self.status_var.set("Terjadi kesalahan.")
                    messagebox.showerror("Pembuat Barcode", ev[1])
        except queue.Empty:
            pass
        self.after(100, self.poll_events)


class CustDropdown:
    """The Cust box's list: typing in the box searches the customer list, the arrow shows all of it.
    Click a name to use it, its ✕ to take it off the list (its SO files stay), '+ Tambah cust' to add
    the name typed in the box."""

    MAX_SHOWN = 12

    def __init__(self, tab, entry, var, arrow):
        self.tab, self.entry, self.var = tab, entry, var
        self.popup = None
        self.hide_job = None
        self.search = True
        entry.bind("<KeyRelease>", self.on_key, add="+")
        entry.bind("<FocusOut>", lambda e: self.schedule_hide(), add="+")
        # the list opens as soon as the box is clicked/entered, even before anything is typed
        entry.bind("<FocusIn>", lambda e: self.show(search=True), add="+")
        entry.bind("<Button-1>", lambda e: self.entry.after(10, lambda: self.show(search=True)), add="+")
        entry.bind("<Escape>", lambda e: self.hide(), add="+")
        entry.bind("<Down>", lambda e: self.show(search=False), add="+")
        arrow.bind("<Button-1>", lambda e: self.toggle())

    def on_key(self, event):
        if event.keysym not in ("Return", "KP_Enter", "Escape", "Tab", "Down", "Up"):
            self.show(search=True)

    def toggle(self):
        if self.popup:
            self.hide()
        else:
            self.entry.focus_set()
            self.show(search=False)

    def schedule_hide(self):   # a click in the list takes the focus first: wait a moment before closing
        self.cancel_hide()
        self.hide_job = self.entry.after(200, self.hide)

    def cancel_hide(self):
        if self.hide_job:
            self.entry.after_cancel(self.hide_job)
            self.hide_job = None

    def hide(self):
        self.cancel_hide()
        if self.popup:
            self.popup.destroy()
            self.popup = None

    def show(self, search=True):
        self.cancel_hide()
        self.search = search
        text = self.var.get().strip()
        names = sorted(self.tab.read_custs(), key=str.lower)
        shown = [n for n in names if text.lower() in n.lower()] if search and text else names
        if self.popup:
            self.popup.destroy()
        self.popup = top = tk.Toplevel(self.entry)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        frame = tk.Frame(top, bg="white", highlightthickness=1, highlightbackground="#888")
        frame.pack(fill="both", expand=True)
        font = ("Segoe UI", 10)
        for name in shown[:self.MAX_SHOWN]:
            row = tk.Frame(frame, bg="white")
            row.pack(fill="x")
            label = tk.Label(row, text=name, bg="white", anchor="w", font=font, padx=6, pady=2, cursor="hand2")
            label.pack(side="left", fill="x", expand=True)
            remove = tk.Label(row, text="✕", bg="white", fg="#c62828", font=font, padx=8, cursor="hand2")
            remove.pack(side="right")
            label.bind("<Button-1>", lambda e, n=name: self.pick(n))
            remove.bind("<Button-1>", lambda e, n=name: self.remove(n))
            parts = (row, label, remove)
            for w in parts:
                w.bind("<Enter>", lambda e, ws=parts: [x.config(bg="#dbe9f9") for x in ws])
                w.bind("<Leave>", lambda e, ws=parts: [x.config(bg="white") for x in ws])
        grey = dict(bg="white", fg="#888", font=("Segoe UI", 9), anchor="w", padx=6)
        if len(shown) > self.MAX_SHOWN:
            tk.Label(frame, text=f"... {len(shown) - self.MAX_SHOWN} lagi - ketik untuk mencari", **grey).pack(fill="x")
        if not shown:
            tk.Label(frame, text="(tidak ada cust yang cocok)" if names else "(daftar cust masih kosong)",
                     **grey).pack(fill="x")
        ttk.Separator(frame).pack(fill="x")
        known = text.lower() in {n.lower() for n in names}
        add = tk.Label(frame, text=f'+ Tambah cust "{text}"' if text and not known else "+ Tambah cust",
                       bg="white", fg="#1f5fa8", font=("Segoe UI", 10, "bold"), anchor="w", padx=6, pady=3,
                       cursor="hand2")
        add.pack(fill="x")
        add.bind("<Button-1>", lambda e: self.add())
        top.update_idletasks()
        x, y = self.entry.winfo_rootx(), self.entry.winfo_rooty() + self.entry.winfo_height()
        top.geometry(f"{max(self.entry.winfo_width() + 24, 260)}x{top.winfo_reqheight()}+{x}+{y}")

    def keep_open(self):
        """After a click that changes the list (✕): back to the box, list still open."""
        def back():
            self.entry.focus_set()
            self.show(search=self.search)
        self.entry.after(10, back)

    def pick(self, name):
        self.var.set(name)
        self.hide()
        self.tab.scan_entry.focus_set()

    def remove(self, name):
        self.tab.remove_cust(name)
        self.tab.show_msg(f"'{name}' dihapus dari daftar cust (file SO-nya tetap ada).")
        self.keep_open()

    def add(self):
        text = self.var.get().strip()
        if not name_part(text):
            self.tab.show_msg("Ketik nama cust baru di kotak Cust, lalu klik '+ Tambah cust'.", ok=False)
            self.keep_open()
            return
        if self.tab.add_cust(text):
            self.tab.show_msg(f"✓  Cust '{text}' ditambahkan ke daftar.")
        else:
            self.tab.show_msg(f"'{text}' sudah ada di daftar cust.", ok=False)
        self.hide()
        self.tab.scan_entry.focus_set()


class ItemDropdown:
    """The scan box's list for typing by hand: shows the Master Barang items whose Kode or Nama
    contain what is typed. Click an item (or pick it with ↓/↑ and Enter) to add it to the SO.
    It waits a moment before opening, so a scanner (which types fast and ends with Enter)
    never sees it."""

    MAX_SHOWN = 12
    DELAY_MS = 250

    def __init__(self, tab, entry, var):
        self.tab, self.entry, self.var = tab, entry, var
        self.popup = None
        self.rows = []          # (item index, row widgets) of the items shown
        self.selected = -1      # row picked with the arrow keys, -1 = none
        self.show_job = None
        self.hide_job = None
        self.lower_cache = (None, [])  # (items list it was made from, [(kode lower, nama lower)])
        entry.bind("<KeyRelease>", self.on_key, add="+")
        entry.bind("<FocusOut>", lambda e: self.schedule_hide(), add="+")
        entry.bind("<Escape>", lambda e: self.hide(), add="+")
        entry.bind("<Down>", lambda e: self.move(1), add="+")
        entry.bind("<Up>", lambda e: self.move(-1), add="+")

    def on_key(self, event):
        if event.keysym in ("Return", "KP_Enter", "Escape", "Tab", "Down", "Up"):
            return
        self.cancel_show()
        self.show_job = self.entry.after(self.DELAY_MS, self.show)

    def cancel_show(self):
        if self.show_job:
            self.entry.after_cancel(self.show_job)
            self.show_job = None

    def schedule_hide(self):   # a click in the list takes the focus first: wait a moment before closing
        if self.hide_job:
            self.entry.after_cancel(self.hide_job)
        self.hide_job = self.entry.after(200, self.hide)

    def hide(self):
        self.cancel_show()
        if self.hide_job:
            self.entry.after_cancel(self.hide_job)
            self.hide_job = None
        if self.popup:
            self.popup.destroy()
            self.popup = None
        self.rows, self.selected = [], -1

    def matches(self, text):
        """Item indexes for `text`: Kode starting with it first, then Kode containing it, then Nama."""
        items = self.tab.app.items
        if self.lower_cache[0] is not items:
            self.lower_cache = (items, [(str(k).lower(), str(n).lower()) for k, n, _h, _s in items])
        t = text.lower()
        starts, in_kode, in_nama = [], [], []
        for i, (k, n) in enumerate(self.lower_cache[1]):
            if k.startswith(t):
                starts.append(i)
            elif t in k:
                in_kode.append(i)
            elif t in n:
                in_nama.append(i)
        return starts + in_kode + in_nama

    def show(self):
        self.show_job = None
        text = self.var.get().strip()
        if not text or not self.tab.app.items:
            self.hide()
            return
        found = self.matches(text)
        if self.popup:
            self.popup.destroy()
        self.rows, self.selected = [], -1
        self.popup = top = tk.Toplevel(self.entry)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        frame = tk.Frame(top, bg="white", highlightthickness=1, highlightbackground="#888")
        frame.pack(fill="both", expand=True)
        for idx in found[:self.MAX_SHOWN]:
            kode, nama = self.tab.app.items[idx][:2]
            row = tk.Frame(frame, bg="white")
            row.pack(fill="x")
            k = tk.Label(row, text=kode, bg="white", anchor="w", font=("Consolas", 11, "bold"),
                         width=16, padx=6, pady=2, cursor="hand2")
            k.pack(side="left")
            n = tk.Label(row, text=nama, bg="white", anchor="w", font=("Segoe UI", 10), padx=4, cursor="hand2")
            n.pack(side="left", fill="x", expand=True)
            parts = (row, k, n)
            pos = len(self.rows)
            for w in parts:
                w.bind("<Button-1>", lambda e, i=idx: self.pick(i))
                w.bind("<Enter>", lambda e, p=pos: self.highlight(p))
            self.rows.append((idx, parts))
        grey = dict(bg="white", fg="#888", font=("Segoe UI", 9), anchor="w", padx=6, pady=2)
        if len(found) > self.MAX_SHOWN:
            tk.Label(frame, text=f"... {len(found) - self.MAX_SHOWN} lagi - ketik lebih lengkap", **grey).pack(fill="x")
        if not found:
            tk.Label(frame, text="(tidak ada barang yang cocok)", **grey).pack(fill="x")
        top.update_idletasks()
        x, y = self.entry.winfo_rootx(), self.entry.winfo_rooty() + self.entry.winfo_height()
        top.geometry(f"{max(self.entry.winfo_width(), 480)}x{top.winfo_reqheight()}+{x}+{y}")

    def highlight(self, pos):
        self.selected = pos
        for p, (_idx, parts) in enumerate(self.rows):
            for w in parts:
                w.config(bg="#dbe9f9" if p == pos else "white")

    def move(self, step):
        if not self.popup:
            self.cancel_show()
            self.show()
        if self.rows:
            self.highlight(max(0, min(len(self.rows) - 1, self.selected + step)))
        return "break"

    def pick_selected(self):
        """Enter in the scan box: if an item is picked with the arrow keys, add it. True if one was."""
        if self.popup and 0 <= self.selected < len(self.rows):
            self.pick(self.rows[self.selected][0])
            return True
        self.hide()
        return False

    def pick(self, idx):
        self.hide()
        self.var.set("")
        self.tab.add_item(idx)
        self.entry.focus_set()


class SOTab(ttk.Frame):
    """The SO tab: scan barcodes into a customer's Sales Order, then print it on the SO form.

    One SO paper is used for the same customer many times, so each customer has one file,
    SO/<Cust>.xlsx. "Load SO" opens it again to add more items; every row remembers the day it
    was added (Tgl msk dos).
    """

    COLUMNS = [  # (id, heading, width, anchor)
        ("no", "No", 40, "center"),
        ("status", "Dicetak", 60, "center"),
        ("tgl_dos", "Tgl Msk Dos", 85, "center"),
        ("kode", "Kode", 120, "w"),
        ("nama", "Nama Barang", 240, "w"),
        ("qty", "Qty", 50, "center"),
        ("sat", "Sat", 50, "center"),
        ("harga", "Harga", 90, "e"),
        ("jumlah", "Jumlah", 100, "e"),
        ("ket", "Catatan", 130, "w"),  # extra note, printed after the Nama Barang in Keterangan
    ]
    EDITABLE = ("qty", "harga", "ket")
    HEADER_FIELDS = [("tgl", "Tgl SO:", 12), ("no_so", "No. SO:", 14), ("cust", "Cust:", 28),
                     ("salesman", "Salesman:", 14)]

    def __init__(self, parent, app):
        super().__init__(parent, padding=12)
        self.app = app
        self.lines = []          # dicts: kode, nama, sat, harga, qty, ket, tgl_dos — in the order added
        self.loaded_path = None  # the customer's SO file this list was loaded from / last saved to
        self.fresh_so = False    # True for a new SO (not loaded from a file); it gets the next free No. SO
        self.op_row = True       # row 8 of the paper is the OP row (new SOs); False for older loaded SOs
        self.editor = None       # the cell being edited, if any
        self.loading = False
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        ttk.Label(self, text="Sales Order (SO)", style="Title.TLabel").grid(row=0, column=0, sticky="w")

        head = ttk.Frame(self)
        head.grid(row=1, column=0, sticky="ew", pady=(6, 10))
        self.head_vars = {}
        for c, (key, label, width) in enumerate(self.HEADER_FIELDS):
            ttk.Label(head, text=label).grid(row=0, column=c * 2, sticky="w", padx=(0 if c == 0 else 16, 4))
            var = tk.StringVar()
            box = ttk.Frame(head)
            box.grid(row=0, column=c * 2 + 1)
            field = ttk.Entry(box, textvariable=var, width=width)
            field.pack(side="left")
            if key == "cust":   # search box + customer list
                arrow = tk.Label(box, text="▼", cursor="hand2", padx=4)
                arrow.pack(side="left")
                self.cust_drop = CustDropdown(self, field, var, arrow)
            field.bind("<Return>", lambda e, k=key: self.on_header_enter(k))
            field.bind("<KP_Enter>", lambda e, k=key: self.on_header_enter(k))
            var.trace_add("write", lambda *_: self.save_draft())
            self.head_vars[key] = var
        self.file_var = tk.StringVar()
        ttk.Label(head, textvariable=self.file_var, foreground="#666").grid(
            row=0, column=len(self.HEADER_FIELDS) * 2, sticky="w", padx=(16, 0))
        self.head_vars["cust"].trace_add("write", lambda *_: self.show_file())
        self.head_vars["no_so"].trace_add("write", lambda *_: self.show_file())
        self.head_vars["tgl"].trace_add("write", lambda *_: self.show_file())

        scan = ttk.Frame(self)
        scan.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(scan, text="Scan barcode:", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.scan_var = tk.StringVar()
        self.scan_entry = ttk.Entry(scan, textvariable=self.scan_var, font=("Consolas", 16), width=22)
        self.scan_entry.pack(side="left", padx=8)
        for key in ("<Return>", "<KP_Enter>", "<Tab>"):  # scanners end a scan with Enter (some with Tab)
            self.scan_entry.bind(key, self.on_scan)
        self.item_drop = ItemDropdown(self, self.scan_entry, self.scan_var)  # list when typing by hand
        self.msg = tk.Label(scan, text="", font=("Segoe UI", 11, "bold"), anchor="w")
        self.msg.pack(side="left", fill="x", expand=True)

        table = ttk.Frame(self)
        table.grid(row=3, column=0, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=[c[0] for c in self.COLUMNS], show="headings",
                                 selectmode="extended")
        for cid, heading, width, anchor in self.COLUMNS:
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, anchor=anchor, stretch=(cid == "nama"))
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd", background="#ececec")
        self.tree.tag_configure("last", background="#fff1a8")  # the item scanned last
        self.tree.tag_configure("printed", foreground="#888888")  # already on the SO paper
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<Key>", self.on_tree_key)
        self.tree.bind("<MouseWheel>", lambda e: self.close_editor())

        ttk.Label(self, text="Barang yang sama di hari yang sama menambah Qty-nya; di hari lain, atau kalau barisnya "
                             "sudah dicetak (abu-abu, ✓), jadi baris baru.  Klik dua kali Qty, Harga atau Catatan untuk "
                             "mengubah (Qty 0 menghapus baris).  'Tambah Barang' menyimpan SO tanpa mencetak; "
                             "'Cetak SO' menyimpan dan mencetak (lain kali hanya baris baru).  Sesudahnya form kosong "
                             "dengan No. SO baru.  Pilih SO di sebelah 'Load SO' (SO cust ini yang belum lunas) untuk menambah ke "
                             f"SO itu.  Satu kertas SO muat {so_page_rows(True)} baris (baris pertama kertas "
                             "baru selalu OP, merah); kalau tidak muat, sisanya dipindah ke "
                             "SO baru.  Centang 'Lunas' kalau SO sudah dibayar (nama file jadi ..._lunas).",
                  foreground="#666", wraplength=980, justify="left").grid(row=4, column=0, sticky="w", pady=(4, 0))

        bottom = ttk.Frame(self)
        bottom.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        self.total_var = tk.StringVar()
        self.lunas_var = tk.BooleanVar()   # paid: the SO's file name gets '_lunas'
        ttk.Checkbutton(bottom, text="Lunas", variable=self.lunas_var, command=self.on_lunas).pack(
            side="right", padx=(8, 0))
        ttk.Button(bottom, text="Cetak SO", style="Big.TButton", command=self.print_so).pack(side="right")
        ttk.Button(bottom, text="Tambah Barang", style="Big.TButton",  # saves the list, doesn't print
                   command=self.save_so).pack(side="right", padx=6)
        self.so_pick_paths = []   # the current customer's SOs that are not lunas yet
        self.so_pick = ttk.Combobox(bottom, state="readonly", width=32, postcommand=self.fill_so_pick)
        self.so_pick.pack(side="right")
        self.so_pick.bind("<<ComboboxSelected>>", self.on_so_pick)
        ttk.Button(bottom, text="Load SO", command=self.load_so).pack(side="right", padx=(6, 4))
        ttk.Button(bottom, text="Hapus terpilih", command=self.remove_selected).pack(side="right", padx=6)
        ttk.Label(bottom, textvariable=self.total_var, style="Summary.TLabel").pack(side="left")

        self.load_draft()
        if self.lines:
            self.refresh()
            self.show_file()
        else:  # nothing unsaved: start with an empty form and the next free No. SO
            self.start_new_so()

    # ---- helpers

    def so_dir(self):
        # the SO tab is built before the "Simpan hasil ke" box exists, so it may not be there yet
        out_var = getattr(self.app, "out_var", None)
        base = out_var.get().strip() if out_var else ""
        # SO files have their own folder, SO_output, next to the sticker output folder
        out = os.path.abspath(base or os.path.join(APP_DIR, "barcode_output"))
        return os.path.join(os.path.dirname(out), "SO_output")

    def target_path(self):
        return so_path(self.so_dir(), self.head_vars["cust"].get(), self.head_vars["no_so"].get().strip(),
                       self.head_vars["tgl"].get(), self.lunas_var.get())

    def show_file(self):
        if self.loading:
            return
        if hasattr(self, "so_pick"):
            self.show_so_pick_hint()
        if not safe_name(self.head_vars["cust"].get()):
            self.file_var.set("File: (isi Cust)")
            return
        path = self.target_path()
        shown = os.path.relpath(path, self.so_dir())
        if self.loaded_path and os.path.normcase(os.path.abspath(self.loaded_path)) == os.path.normcase(path):
            self.file_var.set(f"File: {shown}")
        else:
            self.file_var.set(f"File: {shown}" + ("  (sudah ada)" if os.path.isfile(path) else "  (baru)"))

    @staticmethod
    def today():
        return time.strftime("%d/%m")

    @staticmethod
    def line_key(line):
        return str(line["id"])

    def give_ids(self):
        """Make sure every row has its own id (rows from older drafts or a loaded file may not)."""
        used = {l["id"] for l in self.lines if isinstance(l.get("id"), int)}
        next_id = max(used, default=0) + 1
        for line in self.lines:
            if not isinstance(line.get("id"), int):
                line["id"] = next_id
                next_id += 1
        return next_id

    def line_for(self, key):
        return next((l for l in self.lines if self.line_key(l) == key), None)

    # ---- scanning

    def on_shown(self):
        if self.fresh_so and not self.lines:  # an empty new SO: check the next free number again
            self.head_vars["no_so"].set(str(next_so_number(self.so_dir())))
        self.show_file()
        self.scan_entry.focus_set()

    def show_msg(self, text, ok=True):
        self.msg.config(text=text, fg="#1a7f37" if ok else "#c62828")

    def find_item(self, code):
        """Item index for a scanned code. Also understands scans made with the scanner's Full ASCII
        setting off, e.g. 'UT110-4 /HLSN/I' for UT110-4 (LSN)."""
        for candidate in (code, decode_full_ascii(code)):
            idx = self.app.by_kode.get(candidate.strip().upper())
            if idx is not None:
                return idx
        return None

    def on_scan(self, event=None):
        if self.item_drop.pick_selected():   # an item picked in the typing list with ↓/↑
            return "break"
        code = self.scan_var.get().strip()
        self.scan_var.set("")
        if not code:
            return "break"
        if not self.app.items:
            self.bell()
            self.show_msg("Master Barang masih dimuat - scan lagi sebentar lagi.", ok=False)
            return "break"
        idx = self.find_item(code)
        if idx is None:
            self.bell()
            self.show_msg(f"✗  '{code}' tidak ada di Master Barang", ok=False)
            return "break"
        self.add_item(idx)
        return "break"

    def add_item(self, idx, qty=1):
        """Add Master Barang item `idx` with today's Tgl msk dos. If it was already added today and
        that row is not printed yet, raise its Qty; otherwise (other day, or already on the paper)
        make a new row so nothing that is already printed changes."""
        kode, nama, harga, sat = self.app.items[idx]
        today = self.today()
        line = next((l for l in self.lines if l["kode"] == kode and l.get("tgl_dos") == today
                     and not l.get("printed")), None)
        if line:
            line["qty"] += qty
        else:
            line = {"id": self.give_ids(), "kode": kode, "nama": nama, "sat": sat, "harga": harga,
                    "qty": qty, "ket": "", "tgl_dos": today, "printed": False}
            self.lines.append(line)
        self.refresh(last=self.line_key(line))
        self.show_msg(f"✓  {kode}   {nama}   →  Qty {line['qty']}")
        self.save_draft()
        return line

    def on_header_enter(self, key):
        """Enter in a header box (Cust, Salesman...). A scanner always ends with Enter, so if the box
        now ends with a Kode User, that was a scan into the wrong box: move it to the SO list."""
        if key == "cust":
            self.cust_drop.hide()
        var = self.head_vars[key]
        text = var.get()
        for start in range(len(text)):
            idx = self.find_item(text[start:]) if self.app.items else None
            if idx is not None:
                var.set(text[:start].rstrip())
                self.add_item(idx)
                self.show_msg(f"✓  {self.app.items[idx][0]} masuk ke daftar (tadi ter-scan ke kotak header)")
                break
        self.scan_entry.focus_set()
        return "break"

    def on_tree_key(self, event):
        """If the list has the keyboard focus, send what the scanner types to the scan box."""
        if event.char and event.char.isprintable() and not event.state & CTRL_MASK:
            self.scan_entry.focus_set()
            self.scan_entry.insert("end", event.char)
            return "break"

    # ---- the list

    def refresh(self, last=None):
        self.close_editor(save=False)
        self.tree.delete(*self.tree.get_children())
        first_no = 2 if self.op_row else 1   # same No as on the paper (with the OP row, OP is No 1)
        for n, line in enumerate(self.lines):
            key = self.line_key(line)
            tag = "last" if key == last else ("odd" if n % 2 else "even")
            tags = (tag, "printed") if line.get("printed") else (tag,)
            self.tree.insert("", "end", iid=key, tags=tags, values=(
                n + first_no, "✓" if line.get("printed") else "", line.get("tgl_dos", ""), line["kode"],
                line["nama"], line["qty"], line["sat"],
                rupiah(line["harga"]), rupiah(line["qty"] * line["harga"]), line["ket"]))
        if last and self.tree.exists(last):
            self.tree.see(last)
        qty = sum(l["qty"] for l in self.lines)
        total = sum(l["qty"] * l["harga"] for l in self.lines)
        pages = max(1, math.ceil(len(self.lines) / so_page_rows(self.op_row)))
        new = sum(not l.get("printed") for l in self.lines)
        self.total_var.set(f"{len(self.lines)} baris ({new} belum dicetak)  ·  {qty} pcs  ·  Total {rupiah(total)}"
                           + (f"  ·  {pages} halaman" if pages > 1 else ""))

    def on_double_click(self, event):
        iid, col = self.tree.identify_row(event.y), self.tree.identify_column(event.x)
        if not iid or not col:
            return
        key = self.COLUMNS[int(col[1:]) - 1][0]
        if key not in self.EDITABLE:
            return
        x, y, w, h = self.tree.bbox(iid, col)
        line = self.line_for(iid)
        var = tk.StringVar(value=str(line[key]))
        entry = ttk.Entry(self.tree, textvariable=var, justify="left" if key == "ket" else "center")
        entry.place(x=x, y=y, width=max(w, 80), height=h)
        entry.select_range(0, "end")
        entry.focus_set()
        self.editor = (entry, iid, key, var)
        entry.bind("<Return>", lambda e: self.close_editor())
        entry.bind("<FocusOut>", lambda e: self.close_editor())
        entry.bind("<Escape>", lambda e: self.close_editor(save=False))
        return "break"

    def close_editor(self, save=True):
        if not self.editor:
            return
        entry, line_key, key, var = self.editor
        self.editor = None
        entry.destroy()
        line = self.line_for(line_key)
        if save and line:
            text = var.get().strip()
            if key == "ket":
                line["ket"] = text
            else:
                digits = "".join(ch for ch in text if ch.isdigit())  # accepts "12.600", "Rp 12.600"
                if digits or text == "":
                    value = int(digits) if digits else 0
                    if key == "qty" and value == 0:
                        if line.get("printed"):
                            self.show_msg("Baris yang sudah dicetak tidak bisa dihapus (sudah ada di kertas).",
                                          ok=False)
                        else:
                            self.lines.remove(line)
                    else:
                        line[key] = min(value, 999999999)
            self.refresh(last=None)
            self.save_draft()
            if line.get("printed") and line in self.lines:
                self.show_msg("Catatan: baris ini sudah dicetak - perubahan hanya tersimpan di file, "
                              "tidak di kertas yang sudah dicetak.", ok=False)
        self.scan_entry.focus_set()

    def remove_selected(self):
        selected = set(self.tree.selection())
        if not selected:
            self.show_msg("Klik baris dulu (Ctrl+klik untuk lebih dari satu), lalu Hapus.", ok=False)
            return
        printed = [l for l in self.lines if self.line_key(l) in selected and l.get("printed")]
        if printed:
            self.show_msg(f"{len(printed)} baris sudah dicetak dan tidak bisa dihapus (sudah ada di kertas).",
                          ok=False)
            return
        self.lines = [l for l in self.lines if self.line_key(l) not in selected]
        self.refresh()
        self.save_draft()
        self.show_msg(f"{len(selected)} baris dihapus.")
        self.scan_entry.focus_set()

    # ---- new / load / print

    def start_new_so(self):
        """Empty form for a new SO with the next free No. SO (highest number in any customer folder + 1).
        Salesman stays filled in. Returns the new number."""
        number = next_so_number(self.so_dir())
        self.loading = True
        self.lines = []
        self.loaded_path = None
        self.fresh_so = True
        self.op_row = True
        self.lunas_var.set(False)
        for key, var in self.head_vars.items():
            if key != "salesman":
                var.set("")
        self.head_vars["tgl"].set(time.strftime("%d/%m/%Y"))
        self.head_vars["no_so"].set(str(number))
        self.loading = False
        self.refresh()
        self.save_draft()
        self.show_file()
        self.scan_entry.focus_set()
        return number

    def load_so(self, path=None):
        """Load an SO file (path from the list of open SOs, or chosen in a file window) to add to it."""
        self.close_editor()
        if not path:
            start = self.so_dir() if os.path.isdir(self.so_dir()) else APP_DIR
            folder = os.path.join(start, name_part(self.head_vars["cust"].get()))
            path = filedialog.askopenfilename(title="Load SO (pilih file SO)",
                                              initialdir=folder if os.path.isdir(folder) else start,
                                              filetypes=[("File Excel", "*.xlsx"), ("Semua file", "*.*")])
        if not path:
            return
        if self.lines and not messagebox.askyesno("Load SO", "Daftar sekarang akan diganti dengan SO yang di-load. "
                                                             "Lanjut?"):
            return
        try:
            header, rows = read_so(path)
        except PermissionError:
            messagebox.showerror("Load SO", "File sedang terbuka di Excel. Tutup dulu lalu coba lagi.")
            return
        except Exception as e:
            messagebox.showerror("Load SO", f"Gagal membaca SO:\n\n{e}")
            return
        lines = []
        for row in rows:  # Keterangan holds "Nama Barang - catatan"; split the note back off
            idx = self.app.by_kode.get(row["kode"].upper())
            nama = self.app.items[idx][1] if idx is not None else row["ket"]
            ket = row["ket"]
            note = ket[len(nama):].removeprefix(" - ").strip() if ket.startswith(nama) else ""
            lines.append({"id": len(lines) + 1, "kode": row["kode"], "nama": nama if idx is not None else ket,
                          "sat": row["sat"], "harga": row["harga"], "qty": row["qty"], "ket": note,
                          "tgl_dos": row["tgl_dos"], "printed": row["printed"]})
        self.loading = True
        for key, var in self.head_vars.items():
            var.set(header.get(key, ""))
        if not header.get("cust"):  # older files: use the file name as the Cust
            self.head_vars["cust"].set(os.path.splitext(os.path.basename(path))[0])
        self.lunas_var.set(is_lunas(path))
        self.loading = False
        self.lines = lines
        self.loaded_path = path
        self.fresh_so = False
        self.op_row = header.get("op", False)   # older SOs have no OP row: keep their rows where they are
        self.refresh()
        self.save_draft()
        self.show_file()
        self.show_msg(f"SO '{self.head_vars['cust'].get()}' di-load: {len(lines)} baris" +
                      (" (sudah lunas)" if is_lunas(path) else "") + ". Scan barang untuk menambah.")
        self.scan_entry.focus_set()

    def so_target(self, title):
        """Check the SO can be saved and find its file. Returns (template, folder, file path, file name,
        note) or None if the user has to fix something first. note says if the No. SO had to change."""
        self.close_editor()
        if not self.lines:
            messagebox.showinfo(title, "Scan barang dulu.")
            return None
        if not safe_name(self.head_vars["cust"].get()):
            messagebox.showinfo(title, "Isi nama Cust dulu.")
            return None
        no_so = self.head_vars["no_so"].get().strip()
        if not no_so.isdigit() or int(no_so) < 1:
            messagebox.showinfo(title, "No. SO harus angka (1, 2, 3, ...).")
            return None
        note = ""
        if self.fresh_so:
            # a new SO must not take a number that is already used (e.g. saved meanwhile on another PC)
            used = so_numbers(self.so_dir(), month_year(self.head_vars["tgl"].get()))
            if int(no_so) in used:
                new_no = str(max(used) + 1)
                note = f"No. SO {no_so} sudah dipakai ({', '.join(used[int(no_so)])}) - disimpan sebagai SO {new_no}. "
                self.head_vars["no_so"].set(new_no)
        template = SO_TEMPLATE
        if not os.path.isfile(template):
            template = filedialog.askopenfilename(title="Pilih form SO (Form SO (Sales Order).xlsx)",
                                                  filetypes=[("File Excel", "*.xlsx"), ("Semua file", "*.*")])
            if not template:
                return None
        dst = self.target_path()
        same_file = self.loaded_path and os.path.normcase(os.path.abspath(self.loaded_path)) == \
            os.path.normcase(os.path.abspath(dst))
        if os.path.isfile(dst) and not same_file:
            if not messagebox.askyesno(
                    title, f"File SO ini sudah ada:\n{os.path.relpath(dst, self.so_dir())}\n\n"
                           "Yes = timpa file itu dengan daftar sekarang.\n"
                           "No = batal (pakai 'Load SO' untuk menambah ke SO itu)."):
                return None
        return template, os.path.dirname(dst), dst, os.path.basename(dst), note

    def write_file(self, title, template, out, dst, name, all_printed):
        """Save the whole list to the SO's file. all_printed: mark every row as on the paper."""
        header = {key: var.get() for key, var in self.head_vars.items()}
        lines = [{"kode": l["kode"], "qty": l["qty"], "sat": l["sat"], "harga": l["harga"],
                  "tgl_dos": l.get("tgl_dos", ""), "ket": l["nama"] + (f" - {l['ket']}" if l["ket"] else ""),
                  "printed": all_printed or bool(l.get("printed"))} for l in self.lines]
        try:
            os.makedirs(out, exist_ok=True)
            write_so(template, dst, header, lines, op_row=self.op_row)
        except PermissionError:
            messagebox.showerror(title, f"Tidak bisa menyimpan {name} - file itu sedang terbuka di Excel. "
                                        "Tutup dulu lalu coba lagi.")
            return False
        except Exception as e:
            messagebox.showerror(title, f"Gagal menyimpan SO:\n\n{e}")
            return False
        self.loaded_path = dst
        self.app.output_dir = out
        self.add_cust(header.get("cust", ""))
        return True

    # ---- customers, their open SOs, Lunas

    def cust_file(self):
        return os.path.join(self.so_dir(), "daftar_cust.json")

    def read_custs(self):
        """The customer list, kept in the SO folder so every PC using it sees the same list.
        The first time it is made from the customer folders that already exist."""
        try:
            with open(self.cust_file(), encoding="utf-8") as f:
                return [n for n in json.load(f) if isinstance(n, str) and n.strip()]
        except (OSError, ValueError, TypeError):
            try:
                root = self.so_dir()
                return sorted((d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))), key=str.lower)
            except OSError:
                return []

    def write_custs(self, names):
        try:
            os.makedirs(self.so_dir(), exist_ok=True)
            with open(self.cust_file(), "w", encoding="utf-8") as f:
                json.dump(sorted(names, key=str.lower), f, ensure_ascii=False, indent=1)
        except OSError as e:
            self.show_msg(f"Daftar cust tidak bisa disimpan: {e}", ok=False)

    def add_cust(self, name):
        """Put a customer in the list. False if the name is empty or already there."""
        name = name.strip()
        names = self.read_custs()
        if not name_part(name) or name_part(name).lower() in {name_part(n).lower() for n in names}:
            return False
        self.write_custs(names + [name])
        return True

    def remove_cust(self, name):
        """Take a customer off the list (the SO files are not touched)."""
        self.write_custs([n for n in self.read_custs() if n != name])

    def open_sos(self):
        """(name, path) of the current customer's SO files that are not lunas yet, newest first."""
        cust = name_part(self.head_vars["cust"].get())
        folder = os.path.join(self.so_dir(), cust)
        if not cust or not os.path.isdir(folder):
            return []
        found = []
        for name in os.listdir(folder):
            path = os.path.join(folder, name)
            if name.lower().endswith(".xlsx") and "_SO_" in name and not is_lunas(name) and os.path.isfile(path):
                found.append((os.path.getmtime(path), name[:-5], path))
        return [(name, path) for _, name, path in sorted(found, reverse=True)]

    def fill_so_pick(self):
        sos = self.open_sos()
        self.so_pick_paths = [path for _, path in sos]
        self.so_pick["values"] = [name for name, _ in sos] or ["(tidak ada SO yang belum lunas)"]

    def show_so_pick_hint(self):
        cust = self.head_vars["cust"].get().strip()
        if not cust:
            self.so_pick.set("SO belum lunas (isi Cust dulu)")
        else:
            self.so_pick.set(f"SO {cust} belum lunas: {len(self.open_sos())}")

    def on_so_pick(self, event=None):
        i = self.so_pick.current()
        self.so_pick.selection_clear()
        if 0 <= i < len(self.so_pick_paths):
            self.load_so(self.so_pick_paths[i])
        self.show_so_pick_hint()

    def on_lunas(self):
        """The Lunas box: the buyer has paid (or not). A saved SO is renamed right away (..._lunas.xlsx);
        a new SO gets the name when it is saved."""
        paid = self.lunas_var.get()
        path = self.loaded_path
        if path and os.path.isfile(path) and is_lunas(path) != paid:
            new = lunas_path(path, paid)
            if os.path.exists(new):
                messagebox.showerror("Lunas", f"File ini sudah ada:\n{os.path.basename(new)}")
                self.lunas_var.set(not paid)
                return
            try:
                os.rename(path, new)
            except OSError:
                messagebox.showerror("Lunas", f"Tidak bisa mengganti nama {os.path.basename(path)} - mungkin "
                                              "sedang terbuka di Excel. Tutup dulu lalu coba lagi.")
                self.lunas_var.set(not paid)
                return
            self.loaded_path = new
            self.show_msg(f"✓  {os.path.basename(new)}" + ("  - SO ini lunas." if paid else "  - tanda lunas dihapus."))
        else:
            self.show_msg("SO ini akan disimpan sebagai lunas." if paid else "Tanda lunas dihapus.")
        self.save_draft()
        self.show_file()

    def save_so(self):
        """'Tambah Barang': save the SO without printing anything, then start an empty new SO."""
        target = self.so_target("Tambah Barang")
        if not target:
            return
        template, out, dst, name, note = target
        cust, no_so = self.head_vars["cust"].get(), self.head_vars["no_so"].get()
        room, overflow = self.split_overflow()
        if overflow:  # split now, so the new SO gets its number now and the numbers stay in order
            new_count = sum(not l.get("printed") for l in self.lines)
            what = (f"Kertas SO ini tinggal {room} baris kosong, tapi ada {new_count} baris belum dicetak.\n\n"
                    f"Yes = {room} baris disimpan di SO {no_so} ini, {len(overflow)} baris sisanya "
                    if room else
                    f"Kertas SO ini sudah penuh ({so_page_rows(self.op_row)} baris).\n\nYes = {new_count} baris baru ")
            what += "langsung disimpan sebagai SO baru (Cust sama), supaya No. SO-nya urut.\nNo = batal."
            if not messagebox.askyesno("Tambah Barang - buat SO baru?",
                                       f"Cust: {cust}\nNo. SO: {no_so}\n\n{what}"):
                return
            self.lines = [l for l in self.lines if not any(l is o for o in overflow)]
        if not self.write_file("Tambah Barang", template, out, dst, name, all_printed=False):
            if overflow:  # nothing saved: put the moved rows back
                self.lines += overflow
                self.refresh()
            return
        if overflow:
            waiting = sum(not l.get("printed") for l in self.lines)
            self.continue_in_new_so(cust, overflow, f"✓  {note}Disimpan: {name}" +
                                    (f" - masih {waiting} baris belum dicetak: Load SO {name} lalu "
                                     f"Cetak SO di kertasnya." if waiting else "."),
                                    template, "Tambah Barang")
            return
        new_count = sum(not l.get("printed") for l in self.lines)
        next_no = self.start_new_so()
        self.show_msg(f"✓  {note}Disimpan: {name} ({new_count} baris belum dicetak).  Form baru: SO {next_no}.")

    def print_so(self):
        """'Cetak SO': save the SO and print it, then start an empty new SO. Rows already on the paper are
        not printed again: only the new ones, so the same SO paper can go back into the printer."""
        printed_before = [bool(l.get("printed")) for l in self.lines]
        if self.lines and all(printed_before):
            messagebox.showinfo("Cetak SO", "Semua baris sudah dicetak - tidak ada baris baru.")
            return
        target = self.so_target("Cetak SO")
        if not target:
            return
        template, out, dst, name, note = target
        new_count = printed_before.count(False)
        only_new = new_count < len(self.lines)  # some rows are on the paper already: print just the new ones
        cust, no_so = self.head_vars["cust"].get(), self.head_vars["no_so"].get()
        room, overflow = self.split_overflow()
        if overflow:
            if room:
                what = (f"Kertas SO ini tinggal {room} baris kosong, tapi ada {new_count} baris baru.\n\n"
                        f"Yes = {room} baris dicetak di SO {no_so} ini, {len(overflow)} baris sisanya "
                        f"dipindah ke SO baru (Cust sama) untuk dicetak di kertas SO baru.\n"
                        f"No = batal.")
            else:
                what = (f"Kertas SO ini sudah penuh ({so_page_rows(self.op_row)} baris).\n\n"
                        f"Yes = {new_count} baris baru dipindah ke SO baru (Cust sama) untuk dicetak di "
                        f"kertas SO baru.\nNo = batal.")
            if not messagebox.askyesno("Cetak SO - buat SO baru?", f"Cust: {cust}\nNo. SO: {no_so}\n\n{what}"):
                return
            self.lines = [l for l in self.lines if not any(l is o for o in overflow)]
            printed_before = [bool(l.get("printed")) for l in self.lines]
            new_count = room
        else:
            what = (f"{new_count} baris baru dicetak di kertas SO yang sama." if only_new
                    else f"{new_count} baris dicetak di form SO baru.")
            if not messagebox.askyesno("Cetak SO", f"Apakah sudah benar?\n\nCust: {cust}\nNo. SO: {no_so}\n{what}"):
                return
        if not self.write_file("Cetak SO", template, out, dst, name, all_printed=True):
            if overflow:  # nothing saved: put the moved rows back
                self.lines += overflow
                self.refresh()
            return
        if not new_count:  # the paper was full: nothing to print here, the new rows go to the new SO
            self.continue_in_new_so(cust, overflow, f"{note}{name} disimpan (kertasnya sudah penuh).",
                                    template, "Cetak SO")
            return
        to_print, pages = dst, max(1, math.ceil(len(self.lines) / so_page_rows(self.op_row)))
        if only_new:
            print_dir = os.path.join(out, "cetak_baris_baru")  # kept apart so Load SO doesn't pick them up
            to_print = os.path.join(print_dir, f"{os.path.splitext(name)[0]}_baris_baru_"
                                               f"{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
            try:
                os.makedirs(print_dir, exist_ok=True)
                pages = write_so_new_rows_only(dst, to_print, printed_before)
            except Exception as e:
                messagebox.showerror("Cetak SO", f"Gagal membuat cetakan baris baru:\n\n{e}")
                return
        os.startfile(to_print)
        tip = "tekan Ctrl+P" if pages == 1 else f"tekan Ctrl+P lalu pilih 'Print Entire Workbook' ({pages} halaman)"
        where = "baris baru dibuka di Excel - masukkan kertas SO yang sama" if only_new else "dibuka di Excel"
        if overflow:
            self.continue_in_new_so(cust, overflow, f"{note}{name} disimpan, {where}, {tip}.",
                                    template, "Cetak SO")
            return
        next_no = self.start_new_so()
        self.show_msg(f"{note}{name} disimpan, {where}, {tip}.  Form baru: SO {next_no}.")

    def split_overflow(self):
        """(empty rows left on the SO paper, rows not yet printed that don't fit there). One form has
        so_page_rows() item rows; the printed rows are on the paper already and the unprinted ones go after them."""
        per_page = so_page_rows(self.op_row)
        done = sum(bool(l.get("printed")) for l in self.lines)
        room = (-done) % per_page or (0 if done else per_page)
        return room, [l for l in self.lines if not l.get("printed")][room:]

    def continue_in_new_so(self, cust, lines, done_msg, template, title):
        """The SO paper was full: save the rows that didn't fit as new SOs for the same Cust right away
        (so they get their numbers now, in order). Each new SO is a new paper, so it gets the OP row and
        holds so_page_rows(True) rows. The last one stays in the form."""
        made = []
        per_page = so_page_rows(True)
        while lines:
            chunk, lines = lines[:per_page], lines[per_page:]
            no = self.start_new_so()
            self.head_vars["cust"].set(cust)
            self.lines = chunk
            dst = self.target_path()
            if not self.write_file(title, template, os.path.dirname(dst), dst, os.path.basename(dst),
                                   all_printed=False):
                self.lines = chunk + lines  # not saved: keep the rows in the form
                break
            self.fresh_so = False
            made.append(f"SO {no} ({len(chunk)} baris)")
        self.refresh()
        self.save_draft()
        self.show_file()
        saved = f"SO baru disimpan: {', '.join(made)}" if made else "SO baru belum tersimpan"
        then = ("scan untuk menambah ke SO ini" if title == "Tambah Barang"
                else "masukkan kertas SO baru lalu klik Cetak SO")
        self.show_msg(f"{done_msg}  {saved} - {then}.", ok=False)

    # ---- keep the SO if the app is closed

    def save_draft(self):
        if self.loading:
            return
        try:
            with open(SO_DRAFT_FILE, "w", encoding="utf-8") as f:
                json.dump({"header": {k: v.get() for k, v in self.head_vars.items()}, "lines": self.lines,
                           "loaded_path": self.loaded_path, "fresh_so": self.fresh_so,
                           "op_row": self.op_row, "lunas": self.lunas_var.get()},
                          f, ensure_ascii=False, indent=1)
        except OSError:
            pass

    def draft_op_row(self, draft):
        """Has the SO in the draft the OP row? Drafts from before the OP row existed don't say: then a new
        SO gets it, and a loaded SO keeps what its file has."""
        if "op_row" in draft:
            return bool(draft["op_row"])
        path = draft.get("loaded_path")
        if draft.get("fresh_so") or not path:
            return True
        try:
            return read_so(path)[0]["op"]
        except Exception:
            return False   # file gone or unreadable: don't add an OP row to an older SO

    def load_draft(self):
        self.loading = True
        try:
            with open(SO_DRAFT_FILE, encoding="utf-8") as f:
                draft = json.load(f)
            for key, value in draft.get("header", {}).items():
                if key in self.head_vars:
                    self.head_vars[key].set(value)
            self.lines = [l for l in draft.get("lines", []) if isinstance(l, dict) and "kode" in l]
            for line in self.lines:
                line.setdefault("tgl_dos", self.today())  # drafts saved before Tgl msk dos existed
            self.give_ids()
            self.loaded_path = draft.get("loaded_path")
            self.fresh_so = bool(draft.get("fresh_so"))
            self.op_row = self.draft_op_row(draft)
            self.lunas_var.set(bool(draft.get("lunas")))
        except (OSError, ValueError):
            pass
        if not self.head_vars["tgl"].get():
            self.head_vars["tgl"].set(time.strftime("%d/%m/%Y"))
        self.loading = False


ERROR_LOG = os.path.join(APP_DIR, "barcode_error.log")


def report_crash(error):
    """Started with pythonw there is no console, so write the error to a log and show it."""
    import traceback
    details = "".join(traceback.format_exception(error))
    try:
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write(f"--- {time.strftime('%Y-%m-%d %H:%M:%S')}\n{details}\n")
    except OSError:
        pass
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Pembuat Barcode", f"Aplikasi gagal dibuka:\n\n{error}\n\nDetail disimpan di:\n{ERROR_LOG}")
        root.destroy()
    except Exception:
        pass


if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:
        report_crash(e)
