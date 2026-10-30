"""Starts Pembuat Barcode. Any problem that stops the app from opening - even one that happens
while barcode_app.py itself is being loaded - is shown in a message box and written to
barcode_error.log, because pythonw has no console to show errors in."""

import os
import sys
import time
import traceback

# the folder of this file, or of the .exe when the app is packed as one
HERE = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)


def fail(error):
    details = "".join(traceback.format_exception(error))
    log = os.path.join(HERE, "barcode_error.log")
    try:
        with open(log, "a", encoding="utf-8") as f:
            f.write(f"--- {time.strftime('%Y-%m-%d %H:%M:%S')}  (Python {sys.version.split()[0]}: {sys.executable})\n"
                    f"{details}\n")
    except OSError:
        pass
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Pembuat Barcode", f"Aplikasi gagal dibuka:\n\n{error}\n\n"
                                                f"Python {sys.version.split()[0]}\nDetail disimpan di:\n{log}")
        root.destroy()
    except Exception:
        pass


try:
    import barcode_app
except BaseException as e:  # e.g. a missing library or a mistake in barcode_app.py
    fail(e)
    sys.exit(1)

def selftest(master):
    """`Pembuat Barcode.exe --selftest <Master Barang file>`: check that reading Excel, making sticker
    Word files and SO Excel files all work in this build. Writes selftest_result.txt next to the app."""
    import tempfile
    b = barcode_app
    report = [f"Python {sys.version.split()[0]}  frozen={getattr(sys, 'frozen', False)}",
              f"APP_DIR={b.APP_DIR}", f"SO_TEMPLATE={b.SO_TEMPLATE} exists={os.path.isfile(b.SO_TEMPLATE)}"]
    try:
        items, skipped = b.load_items(master)
        report.append(f"Master Barang: {len(items)} items ({skipped} skipped)")
        tmp = tempfile.mkdtemp()
        stickers = [(b.kode_harga(h, s), k) for k, _, h, s in items[:60]]
        sheets, thin = b.make_sticker_doc(os.path.join(tmp, "stickers.docx"), stickers)
        report.append(f"Stickers: {len(stickers)} on {sheets} sheets, thin={len(thin)}")
        lines = [{"kode": k, "qty": 1, "sat": s, "harga": h, "ket": n, "tgl_dos": "01/01", "printed": i < 3}
                 for i, (k, n, h, s) in enumerate(items[:5])]
        so = os.path.join(tmp, "so.xlsx")
        report.append(f"SO pages: {b.write_so(b.SO_TEMPLATE, so, {'cust': 'TEST', 'no_so': '1'}, lines)}")
        header, back = b.read_so(so)
        report.append(f"SO read back: {len(back)} rows, OP row={header['op']}")
        if len(back) != len(lines) or not header["op"]:
            raise ValueError("SO read back does not match (rows or OP row)")
        report.append(f"SO new rows only: {b.write_so_new_rows_only(so, os.path.join(tmp, 'new.xlsx'), [l['printed'] for l in lines])} page(s)")
        report.append("OK")
    except Exception:
        report.append(traceback.format_exc())
    with open(os.path.join(HERE, "selftest_result.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(report) + "\n")


if "--selftest" in sys.argv:
    selftest(sys.argv[sys.argv.index("--selftest") + 1])
    sys.exit(0)

try:
    barcode_app.App().mainloop()
except Exception as e:
    fail(e)
