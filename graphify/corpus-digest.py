"""Report the current corpus digest and membership, so the publish gate can check a map.

The gate runs in PowerShell and needs two things from the canonical allowlist: the SHA-256
over the allowlisted files as they stand right now, and the list of paths that allowlist
actually selects. Emitting both here keeps `corpus_allowlist.py` the only definition of what
belongs in the public corpus; the gate previously carried its own prefix list, which
accepted the whole of `docs/` and so could drift from the nine approved pages (R2-G03).

Output is a single JSON object on stdout. A non-zero exit or unparseable output must be
treated by the caller as a failure, never as "skip the check" (R2-G02).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from corpus_allowlist import build_manifest  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parent.parent

if __name__ == "__main__":
    manifest = build_manifest(REPO_ROOT)
    print(json.dumps({
        "corpus_sha256": manifest["corpus_sha256"],
        "allowlist_sha256": manifest["allowlist_sha256"],
        "file_count": manifest["file_count"],
        "files": sorted(manifest["files"]),
    }))
