"""Stage the public corpus deterministically and record what it contains (A-01).

Semantic extraction runs over a staging directory rather than the working tree, so what
the map describes depends entirely on what was staged. Doing that by hand meant nobody
could re-create the same corpus later, and nothing stopped a previous map from being
ingested as source on the next run.

This script copies exactly the allowlisted files into a clean staging directory and writes
`corpus-manifest.json` beside them: every file with its SHA-256, plus one digest over the
whole set. `build-public-map.py` re-checks that manifest before it publishes anything, so a
map can always be traced to the bytes it was built from.

Usage:
    python graphify/stage-corpus.py --out <staging-dir>

Then run the extraction inside <staging-dir>, normalize the public extraction before Graphify
constructs the graph, and finally:
    python graphify/prepare-public-extraction.py --corpus <staging-dir>
    python graphify/build-public-map.py --corpus <staging-dir> --spec <extraction-prompt>
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import secrets
import shutil
import stat
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from corpus_allowlist import (  # noqa: E402
    build_manifest,
    is_inside_repository,
    selected_files,
    unlisted_files,
    write_work_inventory,
)

REPO_ROOT = Path(__file__).resolve().parent.parent

# Written into a staging directory once it has been populated. It no longer authorises
# anything — staging deletes nothing at all — but it still identifies a directory as this
# script's output, which is what lets the refusal message say why a target was rejected.
STAGING_MARKER = ".graphify-staging"

# Bumped when the marker's shape changes, so an older marker is not silently accepted.
# Schema 2 binds the marker to the manifest: the nonce is the manifest's run id, so a marker
# and a manifest from different runs no longer describe each other (R4-G01).
MARKER_SCHEMA = 2


def repository_identity() -> str:
    """Identifies this checkout, so a marker written for another one is not honoured."""
    return hashlib.sha256(str(REPO_ROOT).encode("utf-8")).hexdigest()


def read_marker(out_dir: Path) -> dict | None:
    """Reads a staging marker and returns it only when it belongs to this tool and repository.

    <param name="out_dir">Directory that may hold a marker.</param>
    <returns>The marker's fields, or <c>None</c> when it is absent or does not apply.</returns>
    """
    path = out_dir / STAGING_MARKER
    if not path.is_file():
        return None

    try:
        marker = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        # An unreadable or non-JSON marker is not one this tool wrote. Before the contents were
        # checked, any empty file with the right name nominated a directory for deletion (R3-G01).
        return None

    if not isinstance(marker, dict):
        return None
    if marker.get("schema") != MARKER_SCHEMA:
        return None
    if marker.get("tool") != "graphify/stage-corpus.py":
        return None
    if marker.get("repository") != repository_identity():
        return None
    if not isinstance(marker.get("nonce"), str) or len(marker["nonce"]) < 16:
        return None

    return marker


def ensure_safe_staging_target(out_dir: Path) -> None:
    """Refuses any staging target that is not a new directory outside the repository.

    Staging deletes nothing. Earlier versions did, guarded by a `--force` flag and by the
    target's own marker, manifest and inventory — all of which live inside the directory the
    delete would remove, so they could show the contents were self-consistent but never that
    this tool owned them (R7-G02). The delete is gone; what is left is this check plus an
    atomic create, and the diagnostics below explain why a given target is refused.

    <param name="out_dir">Resolved staging directory the caller asked for.</param>
    <exception cref="SystemExit">
        Raised when the target is the repository, lives inside or contains it, or already exists.
    </exception>
    """
    if out_dir == REPO_ROOT:
        raise SystemExit(f"Refusing to stage into the repository itself ({out_dir}).")
    if REPO_ROOT in out_dir.parents:
        raise SystemExit(
            f"Refusing to stage into {out_dir}: it is inside the source tree. A staged corpus "
            "must live outside the repository.")
    if out_dir in REPO_ROOT.parents:
        raise SystemExit(
            f"Refusing to stage into {out_dir}: it contains the source tree.")

    if not out_dir.exists() or not any(out_dir.iterdir()):
        return

    marker = read_marker(out_dir)
    if marker is None:
        raise SystemExit(
            f"{out_dir} is not empty and carries no valid {STAGING_MARKER} marker, so it was not "
            "created by this script for this repository. Refusing to replace it; pick a new "
            "directory.")

    # A marker says the directory was created by this script. It does not say the directory is
    # still only that, and it was possible to drop a valid marker beside someone's own files and
    # have the lot deleted (R4-G01). The manifest is what describes the contents, so it has to be
    # present, has to belong to the same run as the marker, and has to account for every file.
    manifest_path = out_dir / "corpus-manifest.json"
    if not manifest_path.is_file():
        raise SystemExit(
            f"{out_dir} carries a staging marker but no corpus-manifest.json, so nothing describes "
            "what is in it. Refusing to replace it; pick a new directory.")

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        raise SystemExit(
            f"{out_dir} carries a corpus-manifest.json that cannot be read, so nothing describes "
            "what is in it. Refusing to replace it; pick a new directory.") from None

    if not isinstance(manifest, dict) or manifest.get("run_id") != marker["nonce"]:
        raise SystemExit(
            f"The marker and corpus-manifest.json in {out_dir} come from different staging runs, "
            "so neither describes what is there now. Refusing to replace it; pick a new directory.")

    recorded = manifest.get("files")
    if not isinstance(recorded, dict) or not recorded:
        raise SystemExit(
            f"The corpus-manifest.json in {out_dir} lists no files, so it describes nothing. "
            "Refusing to replace it; pick a new directory.")

    # Held to the run the directory says it is, so an inventory carried over from another run
    # accounts for nothing (R4-G01).
    # The digests come with it: a file whose bytes were changed after staging is not the file
    # the manifest describes, however familiar its path looks (R7-G02).
    intruders = unlisted_files(out_dir, recorded, run_id=str(manifest.get("run_id", "")))
    if intruders:
        raise SystemExit(
            f"{out_dir} holds files the corpus manifest does not describe, so it is not only a "
            "staging directory any more:\n  " + "\n  ".join(intruders[:20])
            + "\nRefusing to replace it; pick a new directory.")

    raise SystemExit(
        f"{out_dir} already holds a staged corpus, and nothing outside it can prove those files "
        "are this tool's to delete — the marker, the manifest and the inventory all live inside "
        "the directory. Stage into a new directory instead, and delete the old one yourself if "
        "you know what is in it.")



# A seam for the check-to-use window in the staging copy. Production leaves it None; a fixture
# sets it to swap a name at exactly the moment the containment check has returned and the read has
# not yet happened, which is the only way to reach that window deliberately (R9-G02).
BEFORE_COPY = None


def copy_verified(source, destination, relative):
    """Copies one corpus file, verifying what was opened rather than what was named.

    <para>
        `shutil.copy2(..., follow_symlinks=False)` copies the link itself rather than its target,
        which is safe, but the containment check before it had already been answered against a
        name. Opening the file and comparing the opened object with the name that was checked
        closes the gap between the two on platforms that report inode identity, and narrows it to
        the open call itself everywhere else.
    </para>
    <para>
        O_NOFOLLOW is used where the platform has it. Windows does not, so there the guarantee is
        the identity comparison rather than the open flag; that residual is recorded in the
        threat model rather than described as eliminated.
    </para>

    <param name="source">The file inside the repository.</param>
    <param name="destination">Where in the staging directory it goes.</param>
    <param name="relative">The corpus-relative name, for messages.</param>
    <exception cref="SystemExit">Raised when the name no longer denotes the file that was checked.</exception>
    """
    named = os.lstat(source)
    if stat.S_ISLNK(named.st_mode):
        raise SystemExit(
            f"{relative} is a link, so it names content the repository does not hold. "
            "Refusing to stage it.")

    flags = os.O_RDONLY | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOFOLLOW", 0)
    try:
        handle = os.open(source, flags)
    except OSError as error:
        raise SystemExit(f"Could not read {relative} to stage it: {error}") from None

    try:
        opened = os.fstat(handle)
        if (opened.st_dev, opened.st_ino) != (named.st_dev, named.st_ino):
            raise SystemExit(
                f"{relative} changed between being checked and being read, so what would be "
                "staged is not what was approved. Refusing to stage it.")

        with open(handle, "rb", closefd=False) as reader, open(destination, "wb") as writer:
            shutil.copyfileobj(reader, writer)
    finally:
        os.close(handle)

    shutil.copystat(source, destination, follow_symlinks=False)

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", required=True,
                        help="Staging directory to create; it must not exist yet")
    args = parser.parse_args()

    out_dir = Path(args.out).resolve()
    ensure_safe_staging_target(out_dir)

    # The directory is created here and nowhere else, and its creation is the check: mkdir without
    # exist_ok fails if anything is already there, so a target that appears between the guard
    # returning and this line cannot be adopted. The guard used to be followed by a recursive
    # delete, which that same window turned into an unauthorised rmtree of a directory nobody had
    # approved — and the delete did not even re-read --force (R8-G03).
    try:
        out_dir.mkdir(parents=True)
    except FileExistsError:
        raise SystemExit(
            f"{out_dir} already exists. Staging only ever creates a new directory, so nothing here "
            "is deleted or overwritten; pick a directory that does not exist yet.") from None
    except OSError as error:
        raise SystemExit(f"Could not create the staging directory {out_dir}: {error}") from None

    files = selected_files(REPO_ROOT)
    if not files:
        raise SystemExit("The allowlist selected no files; refusing to stage an empty corpus.")

    for relative in files:
        destination = out_dir / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        source = REPO_ROOT / relative
        # Re-checked at the copy, not only when the list was built: a link planted in between
        # would otherwise be followed out of the repository here (R7-G04).
        if not is_inside_repository(source, REPO_ROOT):
            raise SystemExit(
                f"{relative} does not resolve to a file inside the repository, so it is not "
                "corpus content. Refusing to stage it.")

        # The check above answers for the filesystem as it was a moment ago. Between it and the
        # read there is a window another local process can write in, and the check-then-copy shape
        # is what made that window exploitable rather than merely present (R9-G02). The copy below
        # opens a handle and then asks the handle what it actually opened, so a name swapped in
        # that window is caught rather than followed.
        if BEFORE_COPY is not None:
            BEFORE_COPY(relative, source)

        copy_verified(source, destination, relative)

    # One id for this staging run, carried by both the marker and the manifest so each can be
    # checked against the other (R4-G01).
    run_id = secrets.token_hex(16)

    (out_dir / STAGING_MARKER).write_text(
        json.dumps({
            "schema": MARKER_SCHEMA,
            "tool": "graphify/stage-corpus.py",
            "repository": repository_identity(),
            "nonce": run_id,
            "note": "Staging directory written by graphify/stage-corpus.py. Safe to delete.",
        }, indent=2),
        encoding="utf-8")

    # The manifest describes the staged bytes, not the working tree, so a staged file that
    # was truncated, skipped or edited after copying cannot inherit a digest computed from
    # the repository (R2-G03). Both are compared here so staging fails loudly rather than
    # recording a corpus that does not match the tree it claims to describe.
    manifest = build_manifest(out_dir, run_id)
    live = build_manifest(REPO_ROOT, run_id)
    if manifest["corpus_sha256"] != live["corpus_sha256"]:
        staged_only = sorted(set(manifest["files"]) - set(live["files"]))
        live_only = sorted(set(live["files"]) - set(manifest["files"]))
        changed = sorted(name for name in set(manifest["files"]) & set(live["files"])
                         if manifest["files"][name] != live["files"][name])
        raise SystemExit(
            "The staged copy does not reproduce the working tree, so the corpus cannot be "
            f"trusted.\n  only in staging: {staged_only[:5]}\n  only in tree: {live_only[:5]}\n"
            f"  differing bytes: {changed[:5]}")

    (out_dir / "corpus-manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")

    # The work directory starts with nothing recorded, rather than being exempt by name: a file
    # dropped in there was invisible to the guard while --force still deleted the whole staging
    # root (R4-G01). build-public-map.py records the artifacts it consumes once they exist; nothing
    # else is ever laundered into this list.
    write_work_inventory(out_dir, {}, manifest["run_id"])

    print(f"Staged {manifest['file_count']} files into {out_dir}")
    print(f"  allowlist {manifest['allowlist_sha256'][:12]}")
    print(f"  corpus    {manifest['corpus_sha256'][:12]}")
    print("Run the extraction, prepare-public-extraction.py --corpus it, build the graph, then "
          "build-public-map.py --corpus it.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
