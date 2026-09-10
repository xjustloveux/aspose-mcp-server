"""The single definition of what may enter the public architecture map (G-02, A-01).

Both the staging script and the map generator import this, so the corpus that gets
extracted and the corpus the generator accepts can never drift apart.

Two things are deliberately excluded even though they live under an allowed directory:
`docs/architecture-map/` is the map's own output, and re-ingesting it would let a previous
run's content reappear as source; `docs/assets/` holds a vendored third-party library whose
minified bytes describe nothing about this project.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

# Whole directories whose source is public architecture.
ALLOW_DIRS = ("Core", "Tools", "Handlers", "Results", "Helpers", "Errors")

# Individual repository-root files.
ALLOW_FILES = ("README.md", "Program.cs", "AsposeMcpServer.csproj")

# The published documentation pages, named one by one so a new page has to be approved
# rather than swept in by a directory glob.
ALLOW_DOCS = (
    "docs/index.html",
    "docs/features.html",
    "docs/getting-started.html",
    "docs/configuration.html",
    "docs/tools.html",
    "docs/developers.html",
    "docs/deployment.html",
    "docs/extensions.html",
    "docs/faq.html",
)

# Path fragments that must never appear in the corpus, whatever else matches.
DENY_SUBSTRINGS = (
    "/bin/", "/obj/", "/.git/", "/.claude/", "/graphify-out/", "/coverage-reports/",
    "/sonar-reports/", "/.sonarqube/", "/docs/architecture-map/", "/docs/assets/",
    "/Tests/", "review-backlog", "AGENTS.md", ".lic",
)

# Only these extensions carry source worth extracting.
ALLOW_SUFFIXES = (".cs", ".md", ".html", ".csproj")


# Files a staged corpus carries that are not corpus content: the marker, the manifest itself,
# and the work directory's own inventory.
STAGING_SIDECARS = (".graphify-staging", "corpus-manifest.json")
STAGING_WORK_DIR = "graphify-out"

# Written inside the work directory, naming the artifacts this tooling produced or consumed
# there. Without it the whole subtree was exempt from the guard, so anything could be hidden in it
# and `--force` still deleted the lot (R4-G01).
WORK_INVENTORY = "work-inventory.json"

# The inventory's shape. Bumping it invalidates every older inventory, which is what a change to
# what an entry has to prove requires.
WORK_INVENTORY_SCHEMA = 2


def work_inventory_path(directory: Path) -> Path:
    """Where a staging directory records what its extraction produced.

    <param name="directory">The staging directory.</param>
    <returns>Path of the work directory's inventory file.</returns>
    """
    return directory / STAGING_WORK_DIR / WORK_INVENTORY


def read_work_inventory(directory: Path, run_id: str | None = None) -> set[str]:
    """Reads the artifacts the work directory can prove this tooling produced.

    An entry only counts when the file is still exactly the bytes that were recorded, and when
    the inventory belongs to the staging run being asked about. Anything else — a file nobody
    recorded, a recorded file whose content has since changed, an inventory left over from another
    run — is not accounted for, which is what makes an unknown file stay unknown (R4-G01).

    <param name="directory">The staging directory.</param>
    <param name="run_id">
        The staging run the caller is working in. When given, an inventory written by a different
        run accounts for nothing.
    </param>
    <returns>Directory-relative POSIX paths whose recorded digest still matches the file.</returns>
    """
    path = work_inventory_path(directory)
    if not path.is_file():
        return set()

    try:
        recorded = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return set()

    if not isinstance(recorded, dict):
        return set()
    if recorded.get("schema") != WORK_INVENTORY_SCHEMA:
        return set()
    if run_id is not None and recorded.get("run_id") != run_id:
        return set()

    entries = recorded.get("entries")
    if not isinstance(entries, list):
        return set()

    proven = set()
    for entry in entries:
        if not isinstance(entry, dict):
            continue

        relative = entry.get("path")
        digest = entry.get("sha256")
        if not isinstance(relative, str) or not isinstance(digest, str):
            continue

        target = directory / relative
        if not target.is_file():
            continue
        if hashlib.sha256(target.read_bytes()).hexdigest() != digest:
            continue

        proven.add(relative)

    return proven


def write_work_inventory(directory: Path, artifacts: dict[str, str], run_id: str) -> int:
    """Records the artifacts this tooling produced or consumed in the work directory.

    Only what the caller names is recorded, and each entry carries the bytes it had when it was
    recorded. Walking the directory instead meant a successful publish laundered whatever happened
    to be sitting there into the inventory, after which `--force` would delete it as accounted-for
    content (R4-G01).

    <param name="directory">The staging directory.</param>
    <param name="artifacts">
        Directory-relative POSIX paths the caller produced or consumed, mapped to what each one is.
        A path that does not exist is skipped rather than recorded as something it cannot prove.
    </param>
    <param name="run_id">The staging run these artifacts belong to.</param>
    <returns>How many artifacts were recorded.</returns>
    """
    work_dir = directory / STAGING_WORK_DIR
    work_dir.mkdir(parents=True, exist_ok=True)

    entries = []
    for relative in sorted(artifacts):
        target = directory / relative
        if not target.is_file():
            continue

        entries.append({
            "path": relative,
            "type": artifacts[relative],
            "sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
        })

    work_inventory_path(directory).write_text(
        json.dumps({"schema": WORK_INVENTORY_SCHEMA, "run_id": run_id, "entries": entries},
                   indent=2),
        encoding="utf-8")
    return len(entries)


def unlisted_files(directory: Path, recorded: dict[str, str] | set[str],
                   account_for_work_dir: bool = True,
                   run_id: str | None = None) -> list[str]:
    """Returns every real file in a staged directory that nothing accounts for.

    Corpus content is accounted for by the manifest, and the extraction's own output by the work
    directory's inventory. Exempting the work directory by name instead meant a file dropped in
    there was invisible while `--force` still deleted the whole staging root (R4-G01).

    <param name="directory">The staging directory to walk.</param>
    <param name="recorded">
        What the manifest lists. A mapping of path to digest is checked by content: a file whose
        bytes have changed since it was staged is no longer the file the manifest describes, and
        matching on the path alone let it be edited and still counted as accounted-for (R7-G02).
        A bare set of paths is accepted for callers that have no digests to check against.
    </param>
    <param name="account_for_work_dir">
        Whether the extraction's output has to be accounted for as well. A caller about to delete
        the directory says yes: it may not remove what it cannot explain. A caller about to read a
        few named artifacts out of it says no — it deletes nothing, and the artifacts it does read
        are checked against each other and against the corpus digest.
    </param>
    <param name="run_id">
        The staging run to hold the work inventory to, when it is being accounted for.
    </param>
    <returns>The unaccounted-for paths, sorted, relative to <paramref name="directory" />.</returns>
    """
    produced = read_work_inventory(directory, run_id) if account_for_work_dir else set()
    inventory = work_inventory_path(directory).relative_to(directory).as_posix()
    digests = recorded if isinstance(recorded, dict) else {}

    def accounted(relative: str) -> bool:
        # By exact place, not by file name: a nested `Core/corpus-manifest.json` is somebody's own
        # file, and reading it as the root sidecar exempted it (R7-G02).
        if relative in STAGING_SIDECARS or relative == inventory:
            return True

        if relative in produced:
            return True

        if relative in recorded:
            expected = digests.get(relative)
            if expected is None:
                return True

            target = directory / relative
            return (target.is_file()
                    and hashlib.sha256(target.read_bytes()).hexdigest() == expected)

        return not account_for_work_dir and STAGING_WORK_DIR in Path(relative).parts

    return sorted(
        relative
        for relative in (
            p.relative_to(directory).as_posix()
            for p in directory.rglob("*") if p.is_file())
        if not accounted(relative))


def allowlist_digest() -> str:
    """A digest of the allowlist itself, so a changed policy is visible in the manifest."""
    payload = json.dumps(
        {
            "dirs": list(ALLOW_DIRS),
            "files": list(ALLOW_FILES),
            "docs": list(ALLOW_DOCS),
            "deny": list(DENY_SUBSTRINGS),
            "suffixes": list(ALLOW_SUFFIXES),
        },
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def denied(relative_posix: str) -> bool:
    """Whether a repository-relative POSIX path is excluded by the deny list."""
    probe = "/" + relative_posix.strip("/")
    return any(fragment in probe for fragment in DENY_SUBSTRINGS)


def belongs_to_corpus(relative_posix: str) -> bool:
    """Whether a repository-relative path is one the corpus takes, existing or not.

    <see cref="selected_files" /> answers the same question by walking the tree, so it can only
    answer for files that are still there. A deleted source has to be recognised too — otherwise
    removing a file the map describes leaves nothing to notice it (R4-G03).

    <param name="relative_posix">Repository-relative POSIX path.</param>
    <returns><c>true</c> when the allowlist would select this path.</returns>
    """
    relative = relative_posix.strip("/")
    if not relative or denied(relative):
        return False

    if relative in ALLOW_FILES or relative in ALLOW_DOCS:
        return True

    if not relative.endswith(ALLOW_SUFFIXES):
        return False

    head = relative.split("/", 1)[0]
    return head in ALLOW_DIRS and "/" in relative


def is_linked(path: Path) -> bool:
    """Whether a path is a symbolic link, a junction or any other reparse point.

    <param name="path">The path to test.</param>
    <returns>True when following this name leaves the directory it appears to live in.</returns>
    """
    try:
        return path.is_symlink() or path.is_junction()
    except OSError:
        return True


def is_inside_repository(path: Path, repo_root: Path) -> bool:
    """Whether a path is a real file that is genuinely where its name says it is.

    <para>
        The allowlist is applied to a path's lexical name, while `is_file()`, hashing and copying
        all follow symlinks. A tracked `Core/leak.cs -> /outside/secret` therefore passed the
        allowlist and had its target's bytes staged, extracted and published, with the live and
        staged manifests agreeing because both read the same outside file (R7-G04).
    </para>
    <para>
        Requiring only that the target stay inside the repository was not enough: a junction
        `Core/linked -> Tests/` kept the target inside the repository while moving it out of the
        directory the deny list describes, so `Core/linked/Secret.cs` was accepted and the whole of
        `Tests/` was copied into the public corpus (R8-G02). A name is therefore corpus content
        only when neither it nor any directory above it is a link, and when the path it resolves to
        would itself have been accepted.
    </para>

    <param name="path">The candidate file.</param>
    <param name="repo_root">The repository root the corpus may draw from.</param>
    <returns>True when the file is safe to read as corpus content.</returns>
    """
    try:
        resolved = path.resolve(strict=True)
        root = repo_root.resolve(strict=True)
    except (OSError, RuntimeError):
        return False

    if not resolved.is_file():
        return False

    if not (resolved == root or root in resolved.parents):
        return False

    # Every name between the repository root and the file has to be the directory it appears to
    # be, or the deny list is describing a different place than the one being read.
    current = path if path.is_absolute() else repo_root / path
    while True:
        if is_linked(current):
            return False
        parent = current.parent
        if parent == current or current == root or parent == root:
            break
        current = parent

    # And the place it actually resolves to has to be corpus content in its own right, so a link
    # that survives the checks above still cannot launder a denied directory.
    return not denied(resolved.relative_to(root).as_posix())


def selected_files(repo_root: Path) -> list[str]:
    """Every repository-relative POSIX path that belongs in the corpus, sorted.

    Sorting makes the manifest deterministic: the same tree always produces the same file
    order and therefore the same manifest digest. A path whose resolved target leaves the
    repository is not corpus content, however its name reads (R7-G04).
    """
    chosen: set[str] = set()

    for directory in ALLOW_DIRS:
        base = repo_root / directory
        if not base.is_dir():
            continue
        for path in base.rglob("*"):
            if not path.is_file() or path.suffix not in ALLOW_SUFFIXES:
                continue
            if not is_inside_repository(path, repo_root):
                continue
            relative = path.relative_to(repo_root).as_posix()
            if not denied(relative):
                chosen.add(relative)

    for name in ALLOW_FILES + ALLOW_DOCS:
        path = repo_root / name
        if path.is_file() and is_inside_repository(path, repo_root) and not denied(name):
            chosen.add(Path(name).as_posix())

    return sorted(chosen)


def file_digest(path: Path) -> str:
    """SHA-256 of a file's bytes."""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build_manifest(repo_root: Path, run_id: str | None = None) -> dict:
    """Describes the corpus as content, not as a directory listing.

    The per-file digests are what bind an extraction to the exact sources it read, so a map
    generated from an older tree cannot be presented as describing the current commit.
    """
    files = {
        relative: file_digest(repo_root / relative)
        for relative in selected_files(repo_root)
    }
    combined = hashlib.sha256(
        "\n".join(f"{name}:{digest}" for name, digest in sorted(files.items())).encode("utf-8")
    ).hexdigest()
    return {
        "allowlist_sha256": allowlist_digest(),
        "corpus_sha256": combined,
        # Set by stage-corpus.py; identifies one staging run so the artifacts produced from it
        # can be checked to belong together rather than merely to a valid corpus (R3-G03).
        "run_id": run_id or "",
        "file_count": len(files),
        "allow_dirs": list(ALLOW_DIRS),
        "allow_files": list(ALLOW_FILES),
        "allow_docs": list(ALLOW_DOCS),
        "deny_substrings": list(DENY_SUBSTRINGS),
        "files": files,
    }
