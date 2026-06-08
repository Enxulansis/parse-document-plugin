"""
minerU document parsing helper.
Handles: environment detection, format detection, old-format conversion,
parsing, output reading. All paths use Python native Unicode handling.
"""
import subprocess, sys, os, glob as globmod, shutil, uuid, tempfile

SUPPORTED = (".pdf", ".docx", ".pptx", ".xlsx", ".png", ".jpg", ".jpeg")
LEGACY_MAP = {".doc": ".docx", ".ppt": ".pptx", ".xls": ".xlsx"}


# ---- Environment detection ----

def detect_mineru():
    """Find mineru executable. Tries: conda, then common paths."""
    # Try conda first
    try:
        result = subprocess.run(
            ["conda", "run", "-n", "mineru", "python", "-c",
             "import shutil; print(shutil.which('mineru') or '')"],
            capture_output=True, text=True, timeout=15
        )
        path = result.stdout.strip()
        if path:
            return path
    except Exception:
        pass

    # Try direct conda invocation
    try:
        result = subprocess.run(
            ["conda", "run", "-n", "mineru", "mineru", "--version"],
            capture_output=True, text=True, timeout=15
        )
        if result.returncode == 0:
            # mineru is callable via conda run
            return "conda:mineru"  # special marker
    except Exception:
        pass

    return None


def detect_office_com():
    """Check if Office COM automation is available."""
    try:
        import win32com.client
        import pythoncom
        return True
    except ImportError:
        return False


def check_environment():
    """Run all environment checks. Returns dict of capability flags."""
    env = {
        "mineru": False,
        "mineru_path": None,
        "office_com": False,
        "conda": False,
    }

    # Check conda
    try:
        r = subprocess.run(["conda", "--version"], capture_output=True, text=True, timeout=10)
        env["conda"] = r.returncode == 0
    except Exception:
        pass

    # Check mineru
    mineru_path = detect_mineru()
    if mineru_path:
        env["mineru"] = True
        env["mineru_path"] = mineru_path

    # Check Office COM (Windows only)
    if sys.platform == "win32":
        env["office_com"] = detect_office_com()

    return env


def get_mineru_cmd():
    """Get the command to invoke mineru based on environment."""
    env = check_environment()
    if not env["mineru"]:
        return None, "minerU not found. Install: https://github.com/opendatalab/MinerU"

    path = env["mineru_path"]
    if path == "conda:mineru":
        return ["conda", "run", "-n", "mineru", "mineru"], env
    return [path], env


# ---- Core functions ----

def find_docs(project_dir, filenames=None):
    """Find supported document files. Returns list of absolute paths."""
    if filenames:
        files = []
        for f in filenames:
            full = os.path.join(project_dir, f) if not os.path.isabs(f) else f
            if os.path.exists(full):
                files.append(full)
            else:
                print(f"  [SKIP] Not found: {f}")
        return files
    files = []
    for ext in SUPPORTED:
        files.extend(globmod.glob(os.path.join(project_dir, f"*{ext}")))
    return files


def is_parsed(output_dir):
    """Check if output directory already contains a .md file."""
    if not os.path.exists(output_dir):
        return False
    for root, dirs, files in os.walk(output_dir):
        for f in files:
            if f.endswith(".md"):
                return True
    return False


def get_output_dir(project_dir, filename, is_batch=False):
    """Single: <stem>_<ext>/; Batch: output/<stem>_<ext>/"""
    fname = os.path.basename(filename)
    stem, ext = os.path.splitext(fname)
    dir_name = f"{stem}_{ext.lstrip('.')}"
    if is_batch:
        return os.path.join(project_dir, "output", dir_name)
    return os.path.join(project_dir, dir_name)


def find_cache(project_dir, filename):
    """Search for existing parsed output across single and batch locations."""
    fname = os.path.basename(filename)
    stem, ext = os.path.splitext(fname)
    dir_name = f"{stem}_{ext.lstrip('.')}"
    for d in [os.path.join(project_dir, dir_name),
              os.path.join(project_dir, "output", dir_name)]:
        if is_parsed(d):
            print(f"  [CACHE] Found: {os.path.basename(d)}")
            return d, read_output(d)
    return None, None


def convert_old_format(filepath):
    """Convert legacy formats. Returns new path or None."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in LEGACY_MAP:
        return None

    target_ext = LEGACY_MAP[ext]
    new_path = os.path.splitext(filepath)[0] + target_ext

    if os.path.exists(new_path):
        print(f"  [CONVERT] {new_path} already exists, using it")
        return new_path

    print(f"  [CONVERT] {os.path.basename(filepath)} -> {target_ext}")

    if not detect_office_com():
        print(f"  [CONVERT] Office COM not available.")
        print(f"  [CONVERT] Please manually save as {target_ext}, or run:")
        print(f"           conda run -n mineru pip install pywin32")
        return None

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        if ext == ".doc":
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(filepath)
            doc.SaveAs2(new_path, FileFormat=16)
            doc.Close()
            word.Quit()
        elif ext == ".ppt":
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            pres = ppt.Presentations.Open(filepath)
            pres.SaveAs(new_path, 24)
            pres.Close()
            ppt.Quit()
        elif ext == ".xls":
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(filepath)
            wb.SaveAs(new_path, FileFormat=51)
            wb.Close()
            excel.Quit()

        pythoncom.CoUninitialize()
        print(f"  [CONVERT] Done: {os.path.basename(new_path)}")
        return new_path

    except Exception as e:
        print(f"  [CONVERT] Failed: {e}")
        print(f"  [CONVERT] Please manually save as {target_ext} and try again")
        return None


def parse_document(target, output_dir, mineru_cmd):
    """Run mineru. Handles long filenames. Returns exit code."""
    fname = os.path.basename(target)

    if len(fname) > 80:
        tmp_dir = tempfile.gettempdir()
        short_name = f"_tmp_{uuid.uuid4().hex[:8]}{os.path.splitext(fname)[1]}"
        tmp_path = os.path.join(tmp_dir, short_name)
        shutil.copy2(target, tmp_path)
        print(f"  [LONGNAME] Shortened: {fname[:60]}... -> {short_name}")
        target = tmp_path
        fname = short_name

    print(f"\n  Parsing: {fname}")
    print(f"  Output:  {output_dir}")
    sys.stdout.flush()

    result = subprocess.run(
        mineru_cmd + ["-p", target, "-o", output_dir],
        capture_output=True, text=True
    )
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr)
    status = "OK" if result.returncode == 0 else f"ERR({result.returncode})"
    print(f"  [{status}] {fname}")
    return result.returncode


def read_output(output_dir, max_chars=5000):
    """Find and read .md content from output directory."""
    for root, dirs, files in os.walk(output_dir):
        for f in files:
            if f.endswith(".md"):
                path = os.path.join(root, f)
                with open(path, "r", encoding="utf-8") as fp:
                    content = fp.read()
                if len(content) > max_chars:
                    return content[:max_chars] + f"\n\n... (truncated, total {len(content):,} chars)"
                return content
    return None


def run(project_dir, filenames=None, batch_mode=None):
    """Main entry point. Returns {filename: {status, output_dir, content}}."""
    # Environment check
    mineru_cmd, env = get_mineru_cmd()
    if mineru_cmd is None:
        print("=" * 50)
        print("ERROR: minerU not found.")
        print("Installation: https://github.com/opendatalab/MinerU")
        print("After install, create a conda env: conda create -n mineru")
        print("=" * 50)
        return {}

    # Print env summary
    print(f"minerU: {env.get('mineru_path', 'detected')}")
    print(f"Office COM: {'available' if env.get('office_com') else 'not available (manual conversion needed)'}")
    print(f"Conda: {'available' if env.get('conda') else 'not available'}")

    files = find_docs(project_dir, filenames)
    if not files:
        print("No supported document files found to parse.")
        return {}

    targets = []
    for fp in files:
        ext = os.path.splitext(fp)[1].lower()
        if ext in LEGACY_MAP:
            new_fp = convert_old_format(fp)
            if new_fp:
                targets.append(new_fp)
            else:
                print(f"  [SKIP] Cannot convert: {os.path.basename(fp)}")
        elif ext in SUPPORTED:
            targets.append(fp)
        else:
            print(f"  [SKIP] Unsupported format: {os.path.basename(fp)}")

    is_batch = batch_mode if batch_mode is not None else len(targets) > 1
    print(f"Mode: {'batch' if is_batch else 'single'} ({len(targets)} file(s))")

    results = {}
    for fp in targets:
        filename = os.path.basename(fp)
        output_dir = get_output_dir(project_dir, filename, is_batch=is_batch)

        cache_dir, cache_content = find_cache(project_dir, filename)
        if cache_content is not None:
            print(f"  [CACHE] Using existing parse: {filename}")
            results[filename] = {"status": "skip", "output_dir": cache_dir, "content": cache_content}
            continue

        rc = parse_document(fp, output_dir, mineru_cmd)
        content = read_output(output_dir) if rc == 0 else None
        results[filename] = {
            "status": "ok" if rc == 0 else f"err({rc})",
            "output_dir": output_dir,
            "content": content
        }

    print(f"\n{'='*50}")
    print("Summary:")
    for fname, r in results.items():
        print(f"  [{r['status'].upper():6s}] {fname}")
    print(f"{'='*50}")
    return results
