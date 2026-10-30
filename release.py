"""Make "Pembuat Barcode.exe" and "Pembuat Barcode.zip" in this folder, on any PC with Python.

    python release.py                      build, self-test, zip
    python release.py --master <file.xls>  self-test with this Master Barang file
    python release.py --settings <file>    put this barcode_settings.json in the zip

Steps:
  1. installs the libraries from requirements.txt and PyInstaller if they are missing;
  2. builds the exe (same options as build_exe.bat);
  3. self-test: runs the exe with --selftest on a Master Barang file, if one can be found
     (--master, else "master_path" in barcode_settings.json, else MASTER BARANG UPDATE.xls in this
     folder, else demo\MASTER BARANG DEMO.xlsx, else skipped);
  4. zip: exe + CARA PAKAI.txt + SO form + barcode_settings.json (if found: --settings, else the one
     in this folder), so other PCs get the same paper size.
"""
import importlib.util, json, os, shutil, subprocess, sys, time, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "Pembuat Barcode"
EXE = os.path.join(HERE, NAME + ".exe")
ZIP = os.path.join(HERE, NAME + ".zip")
FORM = os.path.join("SO", "Form SO (Sales Order).xlsx")
MODULES = {"xlrd": "xlrd", "openpyxl": "openpyxl", "PIL": "pillow", "docx": "python-docx",
           "PyInstaller": "pyinstaller"}


def step(text):
    print(f"\n=== {text}", flush=True)


def arg(name):
    return sys.argv[sys.argv.index(name) + 1] if name in sys.argv[:-1] else None


def install_missing():
    missing = [pkg for mod, pkg in MODULES.items() if importlib.util.find_spec(mod) is None]
    if not missing:
        print("semua library sudah ada")
        return
    print("install:", ", ".join(missing))
    subprocess.run([sys.executable, "-m", "pip", "install", *missing], check=True)


def exe_running():
    out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {NAME}.exe"], capture_output=True, text=True).stdout
    return NAME.lower() in out.lower()


def build():
    if exe_running():
        sys.exit(f"GAGAL: {NAME}.exe sedang terbuka - tutup dulu, lalu jalankan lagi.")
    work = os.path.join(HERE, "_build")
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed",
                    "--name", NAME, "--collect-data", "docx", "--hidden-import", "xlrd",
                    "--hidden-import", "openpyxl", "--add-data", f"{os.path.join(HERE, FORM)};SO",
                    "--distpath", os.path.join(work, "dist"), "--workpath", os.path.join(work, "work"),
                    "--specpath", work, os.path.join(HERE, "start_app.pyw")], check=True, cwd=HERE)
    shutil.copy2(os.path.join(work, "dist", NAME + ".exe"), EXE)
    shutil.rmtree(work, ignore_errors=True)
    print("exe:", EXE)


def find_settings():
    for path in (arg("--settings"), os.path.join(HERE, "barcode_settings.json")):
        if path and os.path.isfile(path):
            return path
    return None


def selftest(settings):
    master = arg("--master")
    if not master and settings:
        try:
            master = json.load(open(settings, encoding="utf-8")).get("master_path")
        except (OSError, ValueError):
            master = None
    if not master:   # where the app itself looks: next to the exe
        master = os.path.join(HERE, "MASTER BARANG UPDATE.xls")
        if not os.path.isfile(master):   # public demo: made-up items
            master = os.path.join(HERE, "demo", "MASTER BARANG DEMO.xlsx")
    if not master or not os.path.isfile(master):
        print("dilewati: tidak ada file Master Barang (pakai --master <file.xls>)")
        return None
    result = os.path.join(HERE, "selftest_result.txt")
    if os.path.exists(result):
        os.remove(result)
    subprocess.Popen([EXE, "--selftest", master])   # the exe is windowed: wait for its result file
    for _ in range(180):
        if os.path.exists(result):
            time.sleep(1)
            break
        time.sleep(1)
    text = open(result, encoding="utf-8").read() if os.path.exists(result) else "(tidak ada hasil)"
    os.remove(result) if os.path.exists(result) else None
    print(text.strip())
    ok = text.strip().endswith("OK")
    if not ok:
        sys.exit("GAGAL: self-test tidak OK - zip tidak dibuat.")
    return ok


def make_zip(settings):
    tmp = ZIP + ".new"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(EXE, NAME + ".exe")
        z.write(os.path.join(HERE, "CARA PAKAI.txt"), "CARA PAKAI.txt")
        z.write(os.path.join(HERE, FORM), FORM.replace(os.sep, "/"))
        if settings:
            z.write(settings, "barcode_settings.json")
    os.replace(tmp, ZIP)
    with zipfile.ZipFile(ZIP) as z:
        bad = z.testzip()
        names = z.namelist()
    if bad:
        sys.exit(f"GAGAL: zip rusak ({bad})")
    print("zip:", ZIP, f"({os.path.getsize(ZIP) / 1e6:.1f} MB)")
    for n in names:
        print("   ", n)


if __name__ == "__main__":
    if sys.version_info < (3, 10):
        sys.exit("Perlu Python 3.10 atau lebih baru (di PC utama: Python 3.14).")
    print("Python", sys.version.split()[0], sys.executable)
    step("1. library")
    install_missing()
    step("2. build exe")
    build()
    settings = find_settings()
    step("3. self-test")
    tested = selftest(settings)
    step("4. zip")
    if not settings:
        print("catatan: tidak ada barcode_settings.json - PC lain memakai ukuran kertas standar dari program")
    make_zip(settings)
    print("\nSELESAI" + ("" if tested else " (tanpa self-test)"))
