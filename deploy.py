"""
Deploy schedule_comparison/ to Hugging Face Spaces.
Run from the repo root:  python deploy.py
Requires: git (must be on PATH)
"""
import os, shutil, subprocess, sys, urllib.request, urllib.error, json
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
HF_TOKEN   = os.environ.get("HF_TOKEN", "")  # set env var or paste token here
HF_USER    = "ahmednasser37"
HF_SPACE   = "p6-compare"
SPACE_ID   = f"{HF_USER}/{HF_SPACE}"
CLONE_URL  = f"https://user:{HF_TOKEN}@huggingface.co/spaces/{SPACE_ID}"
DEST       = Path(__file__).parent / "_hf_deploy"
SRC        = Path(__file__).parent / "schedule_comparison"

EXCLUDES = {".venv", "__pycache__", ".env", "tests", ".deploy"}
EXCLUDE_EXTS = {".pyc", ".log"}

# ── Helpers ───────────────────────────────────────────────────────────────────
def run(cmd, cwd=None, check=True):
    print(f"  $ {' '.join(cmd)}")
    return subprocess.run(cmd, cwd=cwd, check=check,
                          capture_output=False, text=True)

def hf_api(method, path, data=None):
    url = f"https://huggingface.co/api/{path}"
    body = json.dumps(data).encode() if data else None
    req = urllib.request.Request(url, data=body, method=method,
          headers={"Authorization": f"Bearer {HF_TOKEN}",
                   "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, json.loads(r.read())
    except urllib.error.HTTPError as e:
        return e.code, {}

def copy_src_to_dest(src: Path, dest: Path):
    """Copy src/ into dest/, skipping excluded files."""
    copied = 0
    for item in src.rglob("*"):
        # Skip excluded dirs and files
        parts = item.relative_to(src).parts
        if any(p in EXCLUDES for p in parts):
            continue
        if item.suffix in EXCLUDE_EXTS:
            continue
        rel = item.relative_to(src)
        target = dest / rel
        if item.is_dir():
            target.mkdir(parents=True, exist_ok=True)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(item, target)
            copied += 1
    print(f"  Copied {copied} files.")

# ── Step 1: Check / create Space ─────────────────────────────────────────────
print("\n[1/5] Checking HuggingFace Space...")
status, _ = hf_api("GET", f"spaces/{SPACE_ID}")
if status == 200:
    print(f"  Space {SPACE_ID} exists.")
elif status == 404:
    print(f"  Space not found — creating {SPACE_ID}...")
    status2, resp = hf_api("POST", "repos/create",
                            {"type": "space", "name": HF_SPACE,
                             "sdk": "docker", "private": False})
    if status2 not in (200, 201):
        print(f"  ERROR creating space: {status2} {resp}")
        sys.exit(1)
    print("  Space created.")
else:
    print(f"  WARNING: unexpected status {status} — continuing anyway.")

# ── Step 2: Clone Space ───────────────────────────────────────────────────────
print("\n[2/5] Cloning Space repo...")
if DEST.exists():
    shutil.rmtree(DEST)
result = run(["git", "clone", CLONE_URL, str(DEST)], check=False)
if result.returncode != 0:
    # Space might be empty — init fresh
    print("  Clone failed (empty Space?) — initialising fresh repo...")
    DEST.mkdir(parents=True, exist_ok=True)
    run(["git", "init"], cwd=DEST)
    run(["git", "remote", "add", "origin", CLONE_URL], cwd=DEST)

# ── Step 3: Sync files ────────────────────────────────────────────────────────
print("\n[3/5] Syncing files...")
# Remove everything except .git
for item in DEST.iterdir():
    if item.name == ".git":
        continue
    if item.is_dir():
        shutil.rmtree(item)
    else:
        item.unlink()
copy_src_to_dest(SRC, DEST)

# ── Step 4: Commit ────────────────────────────────────────────────────────────
print("\n[4/5] Committing...")
run(["git", "config", "user.email", "deploy@local"], cwd=DEST)
run(["git", "config", "user.name", "Deploy"], cwd=DEST)
run(["git", "add", "-A"], cwd=DEST)
result = run(["git", "diff", "--cached", "--quiet"], cwd=DEST, check=False)
if result.returncode == 0:
    print("  Nothing changed — already up to date.")
    sys.exit(0)
run(["git", "commit", "-m", "Deploy from local"], cwd=DEST)

# ── Step 5: Push ─────────────────────────────────────────────────────────────
print("\n[5/5] Pushing to HuggingFace...")
# Try pushing to main; fall back to master
result = run(["git", "push", "origin", "HEAD:main", "--force"], cwd=DEST, check=False)
if result.returncode != 0:
    run(["git", "push", "origin", "HEAD:main", "--force", "--set-upstream"], cwd=DEST)

print("\n✓ Done! Your app will be live in ~1 minute at:")
print(f"  https://huggingface.co/spaces/{SPACE_ID}\n")
