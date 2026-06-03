# -*- coding: utf-8 -*-
"""Optional, fully-gated post-render steps for run.py.

  * to_drive(files, period) — upload the period's deliverables to the studio Google Drive
    (Shared Drive "ניתוחים וישיבות הנהלה" → דשבורדים/<period>/) via the GWS CLI.
  * to_github(root, period) — stage, commit and push the whole project to the private
    backup repo on every run.

Both NEVER raise and NEVER fail the run: each returns a short status string and, on any
problem (tool missing, not authenticated, offline, upload error), returns a
"… skipped — <hint>" message and lets the run continue. This mirrors the existing
optional-feature gating (the lazy python-docx import in targets.py).

Numbers/business logic are untouched here — this module only moves finished files around.
The test gate builds via the renderers directly, so it never invokes these steps.
"""
import datetime
import json
import os
import shutil
import subprocess

# ---- Drive target -----------------------------------------------------------------
# Shared-Drive folder "ניתוחים וישיבות הנהלה"; reports go under דשבורדים/<period>/.
DRIVE_PARENT = "1LXJWw5i9XJO0UDZ251Ll6ipiokOUYu9H"
DASHBOARDS_FOLDER = "דשבורדים"
FOLDER_MIME = "application/vnd.google-apps.folder"

# ---- GitHub target ----------------------------------------------------------------
GIT_REMOTE = "https://github.com/studio139/smb-dashboard.git"
GIT_BRANCH = "main"


# =============================================================== gws plumbing
def _gws_bin():
    return shutil.which("gws")


def _gws(args, timeout=120):
    """Run a `gws` subcommand. Returns (ok, parsed_json_or_None, raw_text). Never raises.

    Args are passed as a list (no shell), so JSON `--params` need no quote-escaping; build
    them with json.dumps(..., ensure_ascii=True) so Hebrew is \\u-escaped (Windows-argv safe).
    The npm shim is gws.cmd — wrap it through the command processor so it launches headless.
    """
    binpath = _gws_bin()
    if not binpath:
        return (False, None, "gws not found")
    cmd = [binpath] + list(args)
    if binpath.lower().endswith((".cmd", ".bat")):
        cmd = [os.environ.get("COMSPEC", "cmd.exe"), "/c"] + cmd
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=timeout)
    except Exception as e:
        return (False, None, str(e))
    raw = (p.stdout or "")
    return (p.returncode == 0, _json_obj(raw), (raw.strip() or (p.stderr or "").strip()))


def _json_obj(text):
    """Extract the single top-level JSON object from gws stdout (it may be preceded by an
    informational 'Using keyring backend' line)."""
    i, j = text.find("{"), text.rfind("}")
    if i >= 0 and j > i:
        try:
            return json.loads(text[i:j + 1])
        except ValueError:
            return None
    return None


def _params(obj):
    return json.dumps(obj, ensure_ascii=True)


def _q_one(quote_value):
    return str(quote_value).replace("\\", "\\\\").replace("'", "\\'")


# =============================================================== Drive helpers
def _authed():
    ok, data, _ = _gws(["auth", "status"])
    return bool(ok and data and data.get("token_valid"))


def _drive_id_of(folder_id):
    ok, data, _ = _gws(["drive", "files", "get",
                        "--params", _params({"fileId": folder_id, "supportsAllDrives": True,
                                             "fields": "id,driveId"})])
    return data.get("driveId") if (ok and data) else None


def _list(parent, drive_id, only_folders=False, name=None, fields="files(id,name)"):
    q = "'%s' in parents and trashed = false" % parent
    if only_folders:
        q += " and mimeType = '%s'" % FOLDER_MIME
    if name is not None:
        q += " and name = '%s'" % _q_one(name)
    ok, data, _ = _gws(["drive", "files", "list",
                        "--params", _params({"q": q, "corpora": "drive", "driveId": drive_id,
                                             "includeItemsFromAllDrives": True,
                                             "supportsAllDrives": True, "fields": fields})])
    return (data.get("files") or []) if (ok and data) else []


def _ensure_folder(parent, name, drive_id):
    hits = _list(parent, drive_id, only_folders=True, name=name)
    if hits:
        return hits[0]["id"]
    ok, data, _ = _gws(["drive", "files", "create",
                        "--json", _params({"name": name, "mimeType": FOLDER_MIME,
                                          "parents": [parent]}),
                        "--params", _params({"supportsAllDrives": True, "fields": "id,name"})])
    return data.get("id") if (ok and data) else None


def _cwd_relative(path):
    """gws only accepts --upload paths INSIDE its working directory. Returns (arg, temp_or_None):
    the real file's cwd-relative path when it lies under cwd (the normal case — outputs are in
    the project tree); otherwise an ASCII-named temp copy placed directly in cwd, which also
    sidesteps any Hebrew-on-cmdline issue."""
    cwd = os.getcwd()
    try:
        rel = os.path.relpath(path, cwd)
    except ValueError:
        rel = ".."
    if not rel.startswith(".."):
        return (rel, None)
    tmp = os.path.join(cwd, "smbpub_%d%s" % (os.getpid(), os.path.splitext(path)[1] or ".bin"))
    shutil.copyfile(path, tmp)
    return (os.path.basename(tmp), tmp)


def _upload(parent, path, drive_id):
    """Create-or-update one file in `parent`. The Drive NAME rides in \\u-escaped JSON metadata
    (Hebrew-safe); the content is referenced by a cwd-relative path. Updating an existing file
    in place keeps re-runs from piling up duplicates."""
    name = os.path.basename(path)            # the real (possibly Hebrew) Drive file name
    hits = _list(parent, drive_id, name=name)
    arg, tmp = _cwd_relative(path)
    try:
        if hits:  # replace content; the existing Hebrew name is kept
            ok, data, _ = _gws(["drive", "files", "update",
                                "--params", _params({"fileId": hits[0]["id"], "supportsAllDrives": True,
                                                    "fields": "id,name"}),
                                "--upload", arg], timeout=300)
        else:
            ok, data, _ = _gws(["drive", "files", "create",
                                "--json", _params({"name": name, "parents": [parent]}),
                                "--params", _params({"supportsAllDrives": True, "fields": "id,name"}),
                                "--upload", arg], timeout=300)
        return bool(ok and data and data.get("id"))
    finally:
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def _resolve_leaf(segments):
    """Find-or-create דשבורדים/<seg1>/<seg2>/…/ (each level). Returns (parent_id, drive_id, rel);
    parent_id is None when the chain can't be reached/created."""
    drive_id = _drive_id_of(DRIVE_PARENT)
    chain = [DASHBOARDS_FOLDER] + list(segments)        # דשבורדים → year → type → period
    rel = "/".join(chain)
    if not drive_id:
        return None, None, rel
    parent = DRIVE_PARENT
    for seg in chain:
        parent = _ensure_folder(parent, seg, drive_id) if parent else None
    return parent, drive_id, rel


def to_drive(files, segments):
    """Upload `files` into Drive folder דשבורדים/<seg1>/<seg2>/…/ — find-or-create EACH level
    (e.g. segments = [year, type, period]). Existing files of the same name are replaced.
    Verifies by listing the leaf afterward. Returns a one-line status string; never raises."""
    if not _gws_bin():
        return "Drive upload skipped — gws CLI not found"
    if not _authed():
        return "Drive upload skipped — run `gws auth login`"
    try:
        parent, drive_id, rel = _resolve_leaf(segments)
        if not parent:
            return "Drive upload skipped — could not find/create %s" % rel
        sent = sum(1 for p in files if os.path.exists(p) and _upload(parent, p, drive_id))
        names = [f.get("name") for f in _list(parent, drive_id, fields="files(name)")]
        return "Drive: %d/%d uploaded to %s/ — now holds: %s" % (
            sent, len(files), rel, ", ".join(names) if names else "(empty)")
    except Exception as e:
        return "Drive upload skipped — %s" % e


def to_drive_preserve(files, segments):
    """Like to_drive, but NEVER overwrites: upload each file ONLY if a file of that name is not
    already in the leaf folder. Used for the studio-filled targets Word — a doc already on Drive
    is left untouched (mirrors the local existence guard). Returns a status string; never raises."""
    if not _gws_bin():
        return "Drive targets skipped — gws CLI not found"
    if not _authed():
        return "Drive targets skipped — run `gws auth login`"
    try:
        parent, drive_id, rel = _resolve_leaf(segments)
        if not parent:
            return "Drive targets skipped — could not find/create %s" % rel
        done = []
        for p in files:
            if not os.path.exists(p):
                continue
            name = os.path.basename(p)
            if _list(parent, drive_id, name=name):
                done.append(name + " (kept)")             # already on Drive — never overwrite
            elif _upload(parent, p, drive_id):
                done.append(name + " (uploaded)")
            else:
                done.append(name + " (failed)")
        return "Drive targets → %s/: %s" % (rel, ", ".join(done) if done else "(none)")
    except Exception as e:
        return "Drive targets skipped — %s" % e


# =============================================================== GitHub backup
def _git(args, root, timeout=120):
    try:
        p = subprocess.run(["git", "-C", root] + list(args), capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=timeout)
        return (p.returncode, p.stdout or "", p.stderr or "")
    except Exception as e:
        return (1, "", str(e))


def to_github(root, period):
    """Stage everything, commit, and push to origin/main. Returns a status string; gated so
    any failure (offline / auth / nothing-to-commit) is reported and the run continues."""
    if not shutil.which("git"):
        return "GitHub backup skipped — git not found"
    rc, _, _ = _git(["rev-parse", "--is-inside-work-tree"], root)
    if rc != 0:
        return "GitHub backup skipped — not a git repo yet"
    _git(["add", "-A"], root)
    stamp = datetime.date.today().strftime("%Y-%m-%d")
    rc, out, err = _git(["commit", "-m", "auto-backup %s %s" % (period, stamp)], root)
    nothing = "nothing to commit" in (out + err).lower()
    rc, out, err = _git(["push", "origin", GIT_BRANCH], root, timeout=240)
    if rc != 0:
        tail = (err.strip().splitlines() or ["offline/auth"])[-1]
        return "GitHub backup: push failed — %s" % tail
    return "GitHub: nothing new (pushed pending)" if nothing else "GitHub: backed up & pushed (%s)" % period
