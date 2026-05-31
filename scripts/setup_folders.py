"""One-time, idempotent reorganization to the PER-MONTH folder model.

Run:  C:\\Python314\\python.exe scripts\\setup_folders.py [YYYY-MM]
(default month = 2026-04)

Per-month model:
  inputs/YYYY-MM/   <- a snapshot of the 5 cumulative exports for that run; it is
                       ALSO the archive (nothing is copied out afterwards).
  הוצאות קבועות.xlsx stays in the root (updated in place).

Migration performed (each step is skip-safe):
  * inputs/current/*.xlsx      -> inputs/<month>/   (then remove empty current/)
  * loose root *.xlsx exports  -> inputs/<month>/   (space-named expenses file stays)
  * delete legacy archive/      (its contents were copies)
  * delete legacy inputs/meetings/ (auto-emitted blanks; meeting Word now lives in
                                    the quarter-closing month folder)
Ensures references/ scripts/ templates/ outputs/ exist.
"""
import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENSURE = ["references", "scripts", "templates", "outputs", "inputs"]


def _move_into(src, dest_dir, log):
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, os.path.basename(src))
    if os.path.abspath(src) == os.path.abspath(dest):
        return
    if os.path.exists(dest):
        log.append("SKIP (exists): " + os.path.relpath(dest, ROOT))
        return
    shutil.move(src, dest)
    log.append("MOVED: " + os.path.basename(src) + " -> " + os.path.relpath(dest_dir, ROOT))


def migrate(month, log):
    for d in ENSURE:
        os.makedirs(os.path.join(ROOT, d), exist_ok=True)
    month_dir = os.path.join(ROOT, "inputs", month)
    os.makedirs(month_dir, exist_ok=True)

    # 1) inputs/current/*.xlsx -> inputs/<month>/
    cur = os.path.join(ROOT, "inputs", "current")
    if os.path.isdir(cur):
        for f in sorted(os.listdir(cur)):
            if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                _move_into(os.path.join(cur, f), month_dir, log)
        if not os.listdir(cur):
            os.rmdir(cur)
            log.append("removed empty inputs/current")

    # 2) loose root exports (xlsx WITHOUT a space) -> inputs/<month>/
    for f in sorted(os.listdir(ROOT)):
        p = os.path.join(ROOT, f)
        if os.path.isfile(p) and f.lower().endswith(".xlsx") and " " not in f:
            _move_into(p, month_dir, log)

    # 3) delete legacy archive/ (copies)
    arch = os.path.join(ROOT, "archive")
    if os.path.isdir(arch):
        shutil.rmtree(arch)
        log.append("deleted legacy archive/")

    # 4) delete legacy inputs/meetings/
    meet = os.path.join(ROOT, "inputs", "meetings")
    if os.path.isdir(meet):
        shutil.rmtree(meet)
        log.append("deleted legacy inputs/meetings/")


def main():
    month = sys.argv[1] if len(sys.argv) > 1 else "2026-04"
    log = []
    migrate(month, log)
    print("=== setup_folders (per-month) ===")
    for line in log:
        print(" " + line)
    md = os.path.join(ROOT, "inputs", month)
    n = len([f for f in os.listdir(md) if f.lower().endswith(".xlsx")]) if os.path.isdir(md) else 0
    root_xlsx = [f for f in os.listdir(ROOT) if f.lower().endswith(".xlsx")]
    print(f"inputs/{month} xlsx = {n}; root xlsx kept = {len(root_xlsx)}")
    print("done")


if __name__ == "__main__":
    main()
