"""Build the public architecture map from an already-extracted Graphify corpus.

Covers G-03 through G-07: it clusters and labels the graph, records the diagnostics for
both the intermediate extraction and the final graph, and writes the public page, metadata,
and lossless relationship catalogue under docs/architecture-map/.

The extraction itself is not done here. Semantic extraction needs an LLM and explicit
authorisation to send source text to it, so it stays a separate, deliberate step; this
script consumes whatever `<corpus>/graphify-out/.graphify_extract.json` already holds.

Usage:
    python graphify/build-public-map.py --corpus <staging-dir> [--out docs/architecture-map]

The corpus directory must contain a `graphify-out` produced by the pinned Graphify
version recorded in graphify/graphify.version; the script refuses to run on a mismatch
so a map can always be traced back to the toolchain that produced it.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from collections import Counter
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from corpus_allowlist import (  # noqa: E402
    ALLOW_DIRS,
    ALLOW_DOCS,
    ALLOW_FILES,
    belongs_to_corpus,
    build_manifest,
    STAGING_WORK_DIR,
    unlisted_files,
    write_work_inventory,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
VERSION_FILE = REPO_ROOT / "graphify" / "graphify.version"

# Where the Graphify skill records the version actually installed for the current user.
# The pin says what should be there; only this file says what is.
SKILL_VERSION_FILE = Path.home() / ".claude" / "skills" / "graphify" / ".graphify_version"

# Membership is decided by the corpus manifest, which lists the allowlisted files one by
# one. An earlier version compared against a "docs/" prefix, which accepted any page under
# docs/ - including the map's own previous output - while corpus_allowlist.py named only
# nine approved pages. That left three copies of the policy free to drift (R2-G03); the
# manifest is now the only one that decides.


# Diagnostic keys worth publishing. The rest of the diagnostic payload describes the local
# toolchain (interpreter paths, source line numbers inside the Graphify package) and would
# put a machine's directory layout on a public page.
PUBLIC_DIAGNOSTIC_KEYS = (
    "node_count", "unverified_node_count", "raw_edge_count", "non_object_edges",
    "valid_candidate_edges", "missing_endpoint_edges", "dangling_endpoint_edges",
    "self_loop_edges", "exact_duplicate_edges", "directed_unique_endpoint_pairs",
    "directed_same_endpoint_collapsed_edges", "undirected_unique_endpoint_pairs",
    "undirected_same_endpoint_collapsed_edges", "same_endpoint_group_count",
    "relation_variant_groups", "post_build_graph_type", "post_build_node_count",
    "post_build_edge_count", "omitted_external_import_edges",
    "normalized_semantic_file_mentions",
)


def public_diagnostics(raw: dict) -> dict:
    """Keeps only the counts, so no local path or package internal reaches the page."""
    return {k: raw[k] for k in PUBLIC_DIAGNOSTIC_KEYS if k in raw}


def ensure_extraction_health(diagnostics: dict) -> None:
    """Refuse a public extraction whose endpoint integrity was not proved.

    Expected AST imports to unmodelled external namespaces are removed explicitly by
    ``prepare-public-extraction.py`` and counted separately. Every edge left in the extraction
    must have real endpoints; otherwise Graphify silently drops it while building the graph.
    """
    required = (
        "dangling_endpoint_edges", "missing_endpoint_edges", "self_loop_edges",
        "omitted_external_import_edges", "normalized_semantic_file_mentions",
    )
    absent = [key for key in required if key not in diagnostics]
    if absent:
        raise SystemExit(
            "The staged extraction has no complete health evidence (missing "
            + ", ".join(absent)
            + "). Run graphify/prepare-public-extraction.py before graph construction.")

    broken = {key: diagnostics[key] for key in required[:3] if diagnostics.get(key) != 0}
    if broken:
        raise SystemExit(
            "The staged extraction still has unrepresentable relationships: "
            + ", ".join(f"{key}={value}" for key, value in broken.items()))


def read_pinned_versions() -> dict[str, str]:
    """Reads graphify/graphify.version into a dict, ignoring comments and blanks."""
    pinned: dict[str, str] = {}
    for line in VERSION_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        key, _, value = line.partition("=")
        pinned[key.strip()] = value.strip()
    return pinned


def observed_skill_version() -> str:
    """The Graphify skill version actually installed, read from the skill's own marker file."""
    if not SKILL_VERSION_FILE.exists():
        raise SystemExit(
            f"Cannot observe the Graphify skill version: {SKILL_VERSION_FILE} is missing. "
            "The map records provenance, so an unobservable toolchain is not publishable.")
    return SKILL_VERSION_FILE.read_text(encoding="utf-8").strip()


def preflight(pinned: dict[str, str], spec_path: Path) -> dict[str, str]:
    """Fails unless the installed package, skill and extraction prompt match the pin.

    <para>
        Every value returned here is measured on this machine. Writing the pinned value into
        the metadata instead would make the map claim a toolchain nobody verified.
    </para>
    <para>
        The interpreter is this process's own. It used to be read from
        `&lt;corpus&gt;/graphify-out/.graphify_python` and handed to subprocess as argv[0] — before
        the corpus manifest was verified — so any corpus a developer or CI was pointed at could run
        an arbitrary program as the person building the map. Asking a program for its own version is
        not a check: it has already run by the time it answers (R7-G01).
    </para>
    """
    installed = subprocess.run(
        [sys.executable, "-c", "import importlib.metadata as m; print(m.version('graphifyy'))"],
        capture_output=True, text=True, check=True).stdout.strip()
    expected = pinned["package"].split("==")[-1]
    if installed != expected:
        raise SystemExit(
            f"Graphify package mismatch: pinned {expected}, installed {installed}. "
            f"Install the pinned version or raise the pin deliberately.")

    skill = observed_skill_version()
    expected_skill = pinned.get("skill_version", "")
    if skill != expected_skill:
        raise SystemExit(
            f"Graphify skill mismatch: pinned {expected_skill}, installed {skill}. "
            "The skill carries the extraction behaviour, so a map built against a different "
            "one is not reproducible. Run 'graphify install' to update the skill, or re-pin "
            "deliberately and compare node/edge/diagnostic counts.")

    if not spec_path.exists():
        raise SystemExit(f"Extraction prompt not found at {spec_path}.")
    spec_sha = hashlib.sha256(spec_path.read_bytes()).hexdigest()
    if spec_sha != pinned.get("extraction_spec_sha256"):
        raise SystemExit(
            "Extraction prompt has changed since the pin. Re-pin deliberately and compare "
            "node/edge/diagnostic counts before accepting the new output.")

    return {
        "graphify_package": f"graphifyy {installed}",
        "graphify_skill_version_observed": skill,
        "extraction_spec_sha256": spec_sha,
    }


def verify_corpus_manifest(corpus: Path) -> dict:
    """Checks the staged corpus still matches the manifest staging wrote for it.

    Without this the source commit in the metadata is just the current HEAD, which says
    nothing about which bytes were extracted; an old extraction could be published as a
    description of today's tree.
    """
    manifest_path = corpus / "corpus-manifest.json"
    if not manifest_path.exists():
        raise SystemExit(
            f"No corpus-manifest.json in {corpus}. Stage the corpus with "
            "graphify/stage-corpus.py so the extraction can be tied to its sources.")

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    current = build_manifest(REPO_ROOT)
    if manifest.get("allowlist_sha256") != current["allowlist_sha256"]:
        raise SystemExit(
            "The corpus allowlist changed after this corpus was staged. Re-stage and "
            "re-extract so the map matches the policy it claims.")
    if manifest.get("corpus_sha256") != current["corpus_sha256"]:
        raise SystemExit(
            "The working tree no longer matches the staged corpus, so this extraction does "
            "not describe the current source. Re-stage and re-extract.")

    # Matching the working tree is not the same as matching the bytes that were extracted.
    # The manifest used to be computed from the repository after copying, so a staged file
    # that was truncated, deleted or edited afterwards still carried a legitimate digest
    # (R2-G03). Recomputing over the staging directory is what ties the map to what the
    # extractor actually read.
    # Anything present in staging that the manifest does not name. build_manifest only walks
    # what the allowlist selects, so a file added to the staging tree afterwards was invisible to
    # both the digest and the added/missing diff below (R3-G02).
    # The extraction's own output is not asked for an inventory here: this reads a few named
    # artifacts out of the work directory and checks them against each other and against the
    # corpus digest, and it deletes nothing. The inventory is what lets a later --force delete
    # that directory, and it is written below, after the publish succeeds (R4-G01).
    intruders = unlisted_files(corpus, set(manifest.get("files", {})),
                               account_for_work_dir=False,
                               run_id=str(manifest.get("run_id", "")))
    if intruders:
        raise SystemExit(
            "The staged corpus holds files the manifest does not describe, so the extraction "
            "cannot be attributed to a known set of sources:\n  " + "\n  ".join(intruders[:20]))

    staged = build_manifest(corpus)
    if staged["corpus_sha256"] != manifest.get("corpus_sha256"):
        recorded, present = set(manifest.get("files", {})), set(staged["files"])
        missing = sorted(recorded - present)
        added = sorted(present - recorded)
        edited = sorted(name for name in recorded & present
                        if manifest["files"][name] != staged["files"][name])
        raise SystemExit(
            "The staged corpus on disk no longer matches its own manifest, so the extraction "
            "cannot be attributed to these sources.\n"
            f"  missing from staging: {missing[:5]}\n"
            f"  not in the manifest: {added[:5]}\n"
            f"  edited after staging: {edited[:5]}")

    # The run id is what later ties the extraction and the render back to this one staging
    # run rather than merely to a valid corpus (R3-G03), so a manifest without one cannot be
    # published against.
    if len(str(manifest.get("run_id", ""))) != 32:
        raise SystemExit(
            "The corpus manifest carries no staging run id, so the extraction and the rendered "
            "page cannot be tied to it. Re-stage with graphify/stage-corpus.py.")
    return manifest


def porcelain_paths(porcelain: str) -> set[str]:
    """Returns every path a `git status --porcelain -z` run reports as changed.

    A record is `XY<space>path`, and a rename or copy appends its source path as the next
    NUL-separated field. X is the index status and Y the working tree's, so a rename made but not
    staged carries its R in the second column; reading only the first left that source path
    unconsumed, whereupon it was read as the next record and its first three characters sliced off
    as if they were a status field. The source of an unstaged rename was therefore never seen, and
    a mangled path was reported in its place (R4-G03).

    <param name="porcelain">The command's raw NUL-separated output.</param>
    <returns>Repository-relative paths, the sources of renames and copies included.</returns>
    """
    entries = porcelain.split("\0")

    changed: set[str] = set()
    index = 0
    while index < len(entries):
        entry = entries[index]
        index += 1
        if len(entry) < 4:
            continue

        changed.add(entry[3:])
        if (entry[0] in ("R", "C") or entry[1] in ("R", "C")) and index < len(entries):
            changed.add(entries[index])
            index += 1

    return changed


def corpus_is_dirty(changed: set[str], corpus_files: frozenset[str]) -> bool:
    """Says whether any changed path is one of the map's sources.

    Membership is asked of the allowlist, not only of the set of files that happen to exist now: a
    deleted corpus file is absent from the current manifest, so intersecting with it alone reported
    a clean commit for a tree that had lost sources the map describes (R4-G03).

    <param name="changed">Paths the working tree reports as differing from HEAD.</param>
    <param name="corpus_files">Repository-relative paths the map is built from.</param>
    <returns>True when the map's sources are uncommitted.</returns>
    """
    return any(path in corpus_files or belongs_to_corpus(path) for path in changed)


def source_commit(corpus_files: frozenset[str]) -> str:
    """Returns the commit the map describes, marked dirty when its sources are uncommitted.

    <param name="corpus_files">
        Repository-relative paths the map is built from. Only corpus paths decide the mark: any
        other change in the working tree — an untracked note, a scratch script, a test — says
        nothing about the sources this map describes, and marking the map dirty for it made the
        label meaningless (R3-G04).
    </param>
    <returns>The commit SHA, with a "-dirty" suffix when a corpus path differs from it.</returns>
    """
    sha = subprocess.run(["git", "rev-parse", "HEAD"], cwd=REPO_ROOT,
                         capture_output=True, text=True, check=True).stdout.strip()

    # -z rather than the default: git quotes non-ASCII paths otherwise, and this repository has
    # them.
    porcelain = subprocess.run(["git", "status", "--porcelain", "-z"], cwd=REPO_ROOT,
                               capture_output=True, text=True, check=True).stdout
    changed = porcelain_paths(porcelain)

    return f"{sha}-dirty" if corpus_is_dirty(changed, corpus_files) else sha


def to_corpus_relative(source_file: str, corpus_root: Path) -> str:
    """Normalises a source reference to a corpus-relative POSIX path.

    The AST layer already stores relative paths, while semantic extraction records the
    absolute path it was handed. Both describe the same file, so the allowlist check
    compares them in one form.
    """
    normalized = source_file.replace("\\", "/").strip()
    if not normalized:
        return ""
    root = str(corpus_root).replace("\\", "/").rstrip("/") + "/"
    if normalized.lower().startswith(root.lower()):
        normalized = normalized[len(root):]
    return normalized


def approved(source_file: str, corpus_root: Path, allowed: frozenset[str]) -> bool:
    """Whether a node or edge source is one of the files the manifest actually staged.

    <param name="source_file">Source reference recorded by the AST or semantic layer.</param>
    <param name="corpus_root">Staging directory the extraction ran in.</param>
    <param name="allowed">Corpus-relative POSIX paths taken from the corpus manifest.</param>
    <returns>True when the reference names a staged file or carries no source at all.</returns>
    """
    normalized = to_corpus_relative(source_file, corpus_root)
    if not normalized:
        # AST nodes for external symbols (Aspose, BCL) carry no source file.
        return True
    if normalized.startswith("/") or (len(normalized) > 1 and normalized[1] == ":"):
        # Still absolute after relativisation: it came from outside the corpus.
        return False
    return normalized in allowed


# The inline literals the page writes into its script, and that harden_page escapes. One list,
# shared with the verifier's raw-angle check through the metadata it publishes (R21-G04).
HARDENED_LITERALS = frozenset({"RAW_NODES", "RAW_EDGES", "LEGEND", "hyperedges"})


def script_literal(html: str, name: str) -> tuple[int, int]:
    """Locates the JSON literal `const NAME = <array or object>;` by a string-aware balanced scan.

    A regex cannot do this: the non-greedy bracket pattern ended the literal at the first `];` it met, and a
    label is free to contain one. `prefix ]; </script><script>...` in a label ended the capture
    inside the string, so only the prefix was escaped and the verifier, using the same regex,
    checked the same prefix (R22-G01). The scan knows what a JSON string is — double quotes,
    backslash escapes — and counts brackets only outside one.

    <param name="html">The page.</param>
    <param name="name">The constant's name.</param>
    <returns>The start and end offsets of the literal, exclusive of the semicolon.</returns>
    <exception cref="SystemExit">Raised when the literal is missing, unbalanced or unterminated.</exception>
    """
    # A declaration is a statement: it starts a line. The first `const NAME =` anywhere in the
    # page was whatever text came first, and a label in an earlier literal could carry one, so
    # the hardener and the gate escaped a decoy inside a string and left the real literal raw
    # (R23-G01). A JSON string cannot hold a raw line break, so a line-start match cannot be
    # inside one; and there must be exactly one.
    heads = list(re.finditer(r"(?m)^[ \t]*const\s+" + re.escape(name) + r"\s*=\s*", html))
    if not heads:
        raise SystemExit(f"The page carries no {name} declaration at the start of a line.")
    if len(heads) > 1:
        raise SystemExit(f"The page declares {name} {len(heads)} times; a literal must be declared once.")
    start = heads[0].end()
    if start >= len(html) or html[start] not in "[{":
        raise SystemExit(f"The {name} literal does not begin with an array or an object.")

    closers = {"[": "]", "{": "}"}
    stack: list[str] = []
    in_string = False
    escaped = False
    for i in range(start, len(html)):
        c = html[i]
        if in_string:
            if escaped:
                escaped = False
            elif c == "\\":
                escaped = True
            elif c == '"':
                in_string = False
            continue
        if c == '"':
            in_string = True
        elif c in closers:
            stack.append(closers[c])
        elif c in "]}":
            if not stack or stack.pop() != c:
                raise SystemExit(f"The {name} literal closes a bracket it never opened.")
            if not stack:
                end = i + 1
                if not re.match(r"\s*;", html[end:]):
                    raise SystemExit(f"The {name} literal is not terminated by a semicolon.")
                return start, end
    raise SystemExit(f"The {name} literal never closes.")


def ensure_literals_are_disjoint(html: str) -> dict[str, tuple[int, int]]:
    """Locates every hardened literal and refuses a page in which one lies inside another.

    Belt and braces over the line-start anchor: whatever found a declaration, a declaration
    that sits inside another literal's range is text in a string, not a statement (R23-G01).

    <param name="html">The page.</param>
    <returns>Each literal's range by name.</returns>
    <exception cref="SystemExit">Raised when two literals overlap.</exception>
    """
    ranges = {name: script_literal(html, name) for name in sorted(HARDENED_LITERALS)}
    ordered = sorted(ranges.items(), key=lambda item: item[1][0])
    for (first, (_, first_end)), (second, (second_start, _)) in zip(ordered, ordered[1:]):
        if second_start < first_end:
            raise SystemExit(f"The {second} declaration lies inside the {first} literal; refused.")
    return ranges


def literal_value(html: str, name: str):
    """Reads one of the page's literals as data.

    <param name="html">The page.</param>
    <param name="name">The constant's name.</param>
    <returns>The parsed value.</returns>
    <exception cref="SystemExit">Raised when the literal cannot be found or is not JSON.</exception>
    """
    start, end = script_literal(html, name)
    try:
        return json.loads(html[start:end])
    except ValueError as error:
        raise SystemExit(f"The page's {name} literal could not be read: {error}") from None


def script_safe_json(value) -> str:
    """Serialises data for a <script> element.

    Produced by a serializer from parsed data, never by substitution inside text: every
    non-ASCII character (U+2028 and U+2029 among them, which end a JavaScript line) becomes a
    `\\uXXXX` escape, and the three characters that can end the element or start markup —
    `<`, `>`, `&` — become escapes too. Those three never occur outside a JSON string, so the
    replacement cannot touch structure.

    <param name="value">The data.</param>
    <returns>JSON that is safe inside a script element.</returns>
    """
    text = json.dumps(value, ensure_ascii=True, separators=(",", ":"))
    return text.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")


def harden_page(html: str) -> str:
    """Removes the two ways graph text could escape its context in the published page.

    Node labels and community names are derived from source symbols and documentation
    headings, so their content is not under this script's control. Two sinks took them
    without neutralising markup: the graph arrays are written straight into a <script>
    element, where a label containing an end-tag closes it early, and the legend built its
    row with innerHTML and no escaping at all (R2-G05). Both are fixed here rather than by
    trusting the extractor to rewrite the strings.

    <param name="html">The generated page.</param>
    <returns>The page with script-safe data arrays and an escaped legend row.</returns>
    <exception cref="SystemExit">Raised when a sink this is meant to fix is not found.</exception>
    """
    # Every inline JSON literal the template emits, not the two that were first noticed. LEGEND
    # carries community labels — the same source-derived text as the node labels — and a label
    # containing an end-tag closed the script element from inside it just as well (R20-G02).
    # `hyperedges` carries the semantic extraction's ids and labels, and the template always
    # emits it, empty or not (R21-G03 found it).
    # Each literal is located by the balanced scan, read as data and written back by the
    # script-safe serializer; nothing is escaped by substitution inside text (R22-G01). All four
    # are located first and must not overlap (R23-G01).
    ensure_literals_are_disjoint(html)
    for name in sorted(HARDENED_LITERALS):
        start, end = script_literal(html, name)
        html = html[:start] + script_safe_json(literal_value(html, name)) + html[end:]
    ensure_literals_are_disjoint(html)

    # Nothing else may be written into the script as a bare data literal. A new literal added to
    # the template lands here rather than going live unescaped.
    # Any identifier JavaScript allows, not only ALL_CAPS: `const legendData = {...}` is a
    # literal the page would write into its script just as readily (R21-G03). The first run
    # with that widened guard found `const hyperedges = [...]` -- lower-case, source-derived,
    # never hardened -- which is why it is in the set above. What the guard looks at is the
    # right-hand side, not the name: a constant assigned from an expression (`document.…`,
    # `new vis.DataSet(…)`, an arrow function) carries no literal text, and an exactly empty
    # `[]` or `{}` carries none on this page; a non-empty array, object or string literal does.
    literal = r'(?:[\[{]\s*[^\]}\s]|["\'`])'
    unexpected = sorted(set(re.findall(
        r"\bconst\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=\s*" + literal, html)) - HARDENED_LITERALS)
    if unexpected:
        raise SystemExit(
            "The page carries inline data literals this hardening does not know: "
            + ", ".join(unexpected) + ". Add them here before publishing.")

    unescaped = ("${c.label}", "${c.color}", "${c.count}")
    if not all(marker in html for marker in unescaped):
        raise SystemExit(
            "The legend row no longer interpolates the values this hardening escapes; "
            "re-check the template before publishing.")
    html = html.replace("${c.color}", "${esc(c.color)}")
    html = html.replace("${c.label}", "${esc(c.label)}")
    html = html.replace("${c.count}", "${esc(c.count)}")
    return html


def rendered_counts(html: str) -> tuple[int, int]:
    """Counts what the page actually draws, which is not what the graph contains.

    Above five thousand nodes graphify renders a community-aggregated view, so the page
    holds one node per community rather than one per symbol. Publishing only the source
    totals told the reader the page was showing 10,928 nodes when it drew 658 (R2-G04).

    <param name="html">The generated page.</param>
    <returns>The number of nodes and edges the page renders.</returns>
    <exception cref="SystemExit">Raised when the page carries no readable node array.</exception>
    """
    return len(literal_value(html, "RAW_NODES")), len(literal_value(html, "RAW_EDGES"))


def graph_edges(graph: dict) -> list[dict]:
    """Returns a graph's edges, whichever key the export used for them."""
    return graph.get("links") or graph.get("edges") or []


def rendered_nodes(html: str) -> list[dict]:
    """Reads the node array the published page draws.

    <param name="html">The exported page.</param>
    <returns>The drawn nodes.</returns>
    <exception cref="SystemExit">Raised when the page carries no node array.</exception>
    """
    return literal_value(html, "RAW_NODES")


def canonical_source_path(value: object, corpus_root: Path | None) -> str:
    """Reduces a source path to the one spelling both artifacts can be compared in.

    The extraction records where it read a file — an absolute path inside the staging directory —
    while the graph records the repository-relative path. Comparing the two as raw strings would
    refuse every one of the 189 nodes where that difference is the only difference, so the staging
    prefix is removed before they are compared, and nothing else is.

    <param name="value">The recorded path.</param>
    <param name="corpus_root">The staging directory the extraction read from, when known.</param>
    <returns>The path in posix form, relative to the corpus root when it lies inside it.</returns>
    """
    text = str(value or "").replace(chr(92), "/")
    if not text:
        return ""

    path = Path(text)
    if corpus_root is not None and path.is_absolute():
        try:
            return path.resolve().relative_to(corpus_root).as_posix()
        except (ValueError, OSError):
            return path.as_posix()

    return path.as_posix()


def relationship_catalog(extraction: dict, corpus_root: Path,
                         run_id: str) -> dict:
    """Builds the deterministic, lossless public relationship catalogue.

    Graphify deliberately projects its input into a simple directed graph for clustering. When
    several relations have the same ordered endpoints, that projection keeps only one of them.
    The projection is useful for rendering, but it is not a complete record of the extraction.
    This sidecar retains every normalized relationship, including relation variants and repeated
    occurrences, while removing the machine-specific staging prefix from source attributions.

    <param name="extraction">The normalized public extraction.</param>
    <param name="corpus_root">The staged corpus the extraction read.</param>
    <param name="run_id">The manifest run identifier shared by all derived artifacts.</param>
    <returns>A stable JSON-ready catalogue containing every extraction edge.</returns>
    <exception cref="SystemExit">Raised when an extraction edge is not an object.</exception>
    """
    relationships = []
    for index, edge in enumerate(extraction.get("edges", [])):
        if not isinstance(edge, dict):
            raise SystemExit(
                f"Extraction relationship {index} is not an object, so it cannot be published.")
        published = dict(edge)
        if "source_file" in published:
            published["source_file"] = canonical_source_path(
                published.get("source_file"), corpus_root)
        relationships.append(published)

    # Extraction order is not a public contract. Sorting the complete object makes the artifact
    # reproducible without deduplicating it: identical occurrences remain identical occurrences.
    relationships.sort(key=lambda edge: json.dumps(
        edge, sort_keys=True, ensure_ascii=False, separators=(",", ":")))

    # Repeating the same field names 27,000 times pushed an otherwise compact JSON artifact over
    # the public 10 MiB ceiling. Group rows by their exact key set and store values positionally.
    # This is a reversible representation: key-set groups distinguish an absent field from one
    # whose value is null, and repeated rows remain repeated rows.
    grouped: dict[tuple[str, ...], list[list[object]]] = {}
    for edge in relationships:
        fields = tuple(sorted(edge))
        grouped.setdefault(fields, []).append([edge[field] for field in fields])

    return {
        "schema": 1,
        "format": "grouped-fields",
        "corpus_run_id": run_id,
        "relationship_count": len(relationships),
        "groups": [
            {"fields": list(fields), "rows": rows}
            for fields, rows in sorted(grouped.items())
        ],
    }


def catalog_relationships(catalog: dict) -> list[dict]:
    """Expands a grouped relationship catalogue without losing missing/null distinctions."""
    if catalog.get("schema") != 1 or catalog.get("format") != "grouped-fields":
        raise SystemExit("relationships.json has an unsupported schema or storage format.")

    relationships = []
    for group_index, group in enumerate(catalog.get("groups", [])):
        if not isinstance(group, dict):
            raise SystemExit(f"Relationship group {group_index} is not an object.")
        fields = group.get("fields")
        rows = group.get("rows")
        if (not isinstance(fields, list) or not fields
                or any(not isinstance(field, str) or not field for field in fields)
                or len(fields) != len(set(fields)) or not isinstance(rows, list)):
            raise SystemExit(f"Relationship group {group_index} has invalid fields or rows.")
        for row_index, row in enumerate(rows):
            if not isinstance(row, list) or len(row) != len(fields):
                raise SystemExit(
                    f"Relationship group {group_index} row {row_index} does not match its fields.")
            relationships.append(dict(zip(fields, row)))

    if catalog.get("relationship_count") != len(relationships):
        raise SystemExit("relationships.json relationship_count does not match its rows.")
    return relationships


def ensure_relationship_catalog(extraction: dict, catalog: dict, corpus_root: Path,
                                run_id: str) -> None:
    """Refuses a catalogue that drops, invents, or changes an extraction relationship."""
    expected = relationship_catalog(extraction, corpus_root, run_id)
    if catalog != expected:
        raise SystemExit(
            "relationships.json is not the lossless catalogue of this extraction and staging "
            "run; rebuild the public map instead of editing the relationship artifact.")


def label_prefix_is_from_source(prefix: str, source_file: str) -> bool:
    """Whether a label's prefix is the tail of that node's own parent directory path.

    The build disambiguates same-named files by walking *up* from the file, so the prefix it
    produces is always a suffix of the parent path: `Tools/Excel/Properties/PropertiesTool.cs`
    can be labelled `Properties/…`, `Excel/Properties/…` or `Tools/Excel/Properties/…`, and
    nothing else.

    Accepting any prefix ending in a slash let a label be given arbitrary text (R9-G01), and
    accepting any *contiguous run* of segments was still wider than the build: probed directly,
    `Tools` and even the basename `PropertiesTool.cs` were accepted as prefixes of that same
    file, neither of which the builder can emit (§21.5).

    <param name="prefix">The prefix the graph's label carries.</param>
    <param name="source_file">The node's canonical source path.</param>
    <returns>Whether the prefix is a suffix of the file's parent path.</returns>
    """
    if not prefix or not source_file:
        return False

    parent = source_file.split("/")[:-1]
    wanted = prefix.split("/")

    return bool(parent) and len(wanted) <= len(parent) and parent[-len(wanted):] == wanted


def ensure_artifacts_share_one_run(extraction: dict, graph: dict, html: str,
                                   corpus_root: Path | None = None) -> None:
    """Checks the graph derives from the extraction and the page derives from the graph.

    A valid manifest only says the corpus is current; it says nothing about which extraction
    produced the graph, or which graph produced the page. Each was read from disk independently,
    so an older or swapped artifact could be published alongside a fresh manifest (R3-G03).

    Comparing ids alone was not enough: artifacts from a different run of the *same* corpus carry
    the same ids, so an older graph with different labels, different edges and a different
    community partition was accepted as derived from this extraction (R4-G02). What is compared
    here is content — every node, every edge and every drawn community has to be one this
    extraction and this graph actually contain.

    Content comparison was itself too loose to be worth much: measured against the live checker,
    a label could be given an arbitrary prefix, an attribution could be shortened to a substring
    of the real one, a weight could be deleted rather than falsified, an edge could be duplicated
    with the page's aggregate count moved to match, and a dropped `references` edge could be
    covered by any other edge on the pair rather than by the strong relation that actually
    absorbed it. All six were accepted (R9-G01). Every rule below is measured against the real
    artifacts first, so what it admits is what this build actually produces and nothing wider.

    <param name="extraction">The merged AST and semantic extraction.</param>
    <param name="graph">The built graph.</param>
    <param name="html">The exported page.</param>
    <param name="corpus_root">
        The staging directory the extraction read from. Given, source paths are compared after
        the staging prefix is removed; omitted, they are compared as recorded.
    </param>
    <exception cref="SystemExit">Raised when one artifact does not derive from the previous one.</exception>
    """
    # Checked before anything is keyed by id, because keying is what hid it: two extraction nodes
    # sharing an id meant the second silently replaced the first, and a graph carrying only the
    # second was then accepted as derived from the whole extraction. Probed directly against the
    # live checker, that passed (R10-G01). Every comparison below assumes one node per id; that
    # assumption is now verified rather than relied on.
    drawn_page_nodes = rendered_nodes(html)
    for artifact, nodes, allow_identical_copies in (
            ("extraction", extraction.get("nodes", []), True),
            ("graph", graph.get("nodes", []), False),
            ("page", drawn_page_nodes, False)):
        by_id: dict[str, list[dict]] = {}
        blank = 0
        for node in nodes:
            node_id = str(node.get("id") or "")
            if not node_id:
                blank += 1
                continue
            by_id.setdefault(node_id, []).append(node)

        if blank:
            raise SystemExit(
                f"The {artifact} holds {blank} node(s) with no id, so its nodes cannot be "
                "compared with the other artifacts'.")

        if allow_identical_copies:
            # Measured over the real artifacts: the extraction repeats 6 ids, and every one of
            # those repeats is a byte-identical copy — the same symbol found twice, which the
            # dictionary collapses to the same thing either way. Two nodes sharing an id with
            # *different* content are another matter: the second replaced the first before any
            # comparison ran, so a graph carrying only the second passed as derived from the whole
            # extraction (R10-G01).
            offenders = sorted(
                node_id for node_id, group in by_id.items()
                if len({json.dumps(node, sort_keys=True) for node in group}) > 1)
            trouble = "the same id and different content"
        else:
            # The graph and the page are this build's own output, where an id occurs once —
            # measured: zero repeats in either. A repeat there is not a duplicate finding, it is
            # an artifact this check has never seen.
            offenders = sorted(node_id for node_id, group in by_id.items() if len(group) > 1)
            trouble = "the same id"

        if offenders:
            raise SystemExit(
                f"The {artifact} holds more than one node with {trouble}, so one of them would "
                "silently replace the other before anything is compared:"
                + "\n  " + "\n  ".join(offenders[:10]))

    extracted = {str(n.get("id")): n for n in extraction.get("nodes", [])}
    built = {str(n.get("id")): n for n in graph.get("nodes", [])}

    strays = sorted(set(built) - set(extracted))
    if strays:
        raise SystemExit(
            "The graph holds nodes the extraction does not, so it was not built from it:"
            + "\n  " + "\n  ".join(strays[:10]))

    dropped = sorted(set(extracted) - set(built))
    if dropped:
        raise SystemExit(
            "The extraction holds nodes the graph does not, so the graph was built from a "
            "different extraction:" + "\n  " + "\n  ".join(dropped[:10]))

    # The build prefixes path segments onto a label to tell same-named files apart, so the graph's
    # label ends with the extraction's rather than equalling it. Requiring only that it end with
    # "/" + the extraction's label let the prefix be anything at all (R9-G01). Measured over the
    # real artifacts: 11,209 labels are identical and 18 carry a prefix, and every one of those 18
    # prefixes is a run of path segments from that node's own source file — "Excel/DataOperations",
    # "DigitalSignature", "Word/Properties". That is the whole of it, so that is what is accepted.
    def label_agrees(graph_label: str, extraction_label: str, source_file: str) -> bool:
        """Whether the graph's label is the extraction's, with at most a prefix from its path."""
        if graph_label == extraction_label:
            return True
        if not graph_label.endswith("/" + extraction_label):
            return False

        prefix = graph_label[: -len("/" + extraction_label)]
        return label_prefix_is_from_source(prefix, source_file)

    relabelled = sorted(
        node_id for node_id, node in built.items()
        if str(extracted[node_id].get("label") or "")
        and not label_agrees(str(node.get("label") or ""),
                             str(extracted[node_id].get("label") or ""),
                             canonical_source_path(
                                 extracted[node_id].get("source_file")
                                 or node.get("source_file"), corpus_root))
    )
    if relabelled:
        raise SystemExit(
            "The graph labels these nodes differently from the extraction, so the two are from "
            "different runs:" + "\n  " + "\n  ".join(relabelled[:10]))

    # A node whose attribution the extraction knows has to carry it in the graph too. Requiring
    # both sides to be non-empty meant deleting the graph's source_file skipped the comparison
    # altogether, so the provenance could be removed rather than falsified (R8-G04); comparing
    # with `in` then let it be shortened, so a node the extraction attributes to "Core/B.cs" could
    # be published as coming from "B.cs" (R9-G01). Measured after the staging prefix is removed:
    # all 8,783 nodes that carry an attribution on both sides carry the *same* one, so equality is
    # what this build produces.
    misattributed = sorted(
        node_id for node_id, node in built.items()
        if extracted[node_id].get("source_file")
        and (not node.get("source_file")
             or canonical_source_path(node["source_file"], corpus_root)
             != canonical_source_path(extracted[node_id]["source_file"], corpus_root))
    )
    if misattributed:
        raise SystemExit(
            "The graph attributes these nodes to different source files than the extraction, or "
            "drops an attribution the extraction has:"
            + "\n  " + "\n  ".join(misattributed[:10]))

    # Every edge the graph draws has to be one the extraction found, with a relation it gave that
    # pair. A graph from another run over the same nodes fails here even when the ids all match.
    # Keyed by relation *and* confidence: a graph that kept the pair and the relation but
    # relabelled an EXTRACTED edge as INFERRED was accepted, and confidence is what the page
    # renders as certainty (R7-G03). Measured against the real artifacts first — 25,065 graph
    # edges, none of which the extraction had not offered with that exact confidence.
    # Every edge the graph draws has to be one the extraction found, with the relation and the
    # confidence it gave that pair — and no more often than it offered them. Comparing sets lost
    # multiplicity, so one extraction edge could be drawn twice, and the aggregated page then
    # reports the inflated number as the strength of the tie between two communities (R9-G01).
    # Measured against the real artifacts: 25,199 graph edges, every key drawn exactly once, none
    # more often than the extraction offers it.
    def edge_key(edge: dict, source_field: str, target_field: str) -> tuple[str, str, str, str]:
        """The identity of an edge: its direction, its relation and its confidence."""
        return (str(edge.get(source_field)), str(edge.get(target_field)),
                str(edge.get("relation")), str(edge.get("confidence") or "EXTRACTED"))

    offered: Counter = Counter(edge_key(edge, "source", "target")
                               for edge in extraction.get("edges", []))
    drawn_edges: Counter = Counter(edge_key(edge, "source", "target")
                                   for edge in graph_edges(graph))

    relations: dict[tuple[str, str], set[tuple[str, str]]] = {}
    for source, target, relation, confidence in offered:
        relations.setdefault((source, target), set()).add((relation, confidence))

    drawn_by_pair: dict[tuple[str, str], set[tuple[str, str]]] = {}
    for source, target, relation, confidence in drawn_edges:
        drawn_by_pair.setdefault((source, target), set()).add((relation, confidence))

    invented = []
    for (source, target, relation, confidence), count in sorted(drawn_edges.items()):
        available = offered.get((source, target, relation, confidence), 0)
        if not available:
            invented.append(
                f"{source} -> {target} as {relation} [{confidence}] (no such relationship)")
        elif count > available:
            invented.append(
                f"{source} -> {target} as {relation} [{confidence}] drawn {count} times, "
                f"but the extraction offers it {available}")
        if len(invented) >= 10:
            break

    if invented:
        raise SystemExit(
            "The graph holds relationships the extraction does not, so it was built from a "
            "different extraction:" + "\n  " + "\n  ".join(invented))

    # Every edge the graph draws must weigh what this build emits. Nothing looked at weight, so a
    # relationship could be given any prominence the page then renders (R8-G04); then checking it
    # only when the field was present made deleting the weight a way of skipping the check
    # entirely (R9-G01). Measured: all 25,199 edges carry weight 1.0 and none omits it, so a
    # missing weight is as much a graph this check has never seen as a wrong one.
    def weight_disagrees(edge: dict) -> bool:
        """Whether an edge fails to carry the weight every edge of this build carries."""
        if "weight" not in edge or edge.get("weight") is None:
            return True
        try:
            return float(edge["weight"]) != 1.0
        except (TypeError, ValueError):
            return True

    misweighted = sorted(
        f"{edge.get('source')} -> {edge.get('target')} "
        f"(weight {edge.get('weight', 'absent')!r})"
        for edge in graph_edges(graph) if weight_disagrees(edge))
    if misweighted:
        raise SystemExit(
            "The graph weights these relationships differently from every edge this build emits:"
            + "\n  " + "\n  ".join(misweighted[:10]))

    # The build drops an edge only when both endpoints are present and it collapses into a
    # *stronger* edge between the same pair; anything else missing means this graph was not built
    # from this extraction. Checking only the other direction let a graph with no edges at all
    # pass (R4-G02); checking only the pair let a pair keep one relation while another vanished
    # (R8-G04); and requiring merely that *some* edge remain let a `references [INFERRED]` edge be
    # covered by `references [EXTRACTED]`, or by any weak relation at all, so a confidence
    # downgrade could be published as a collapse (R9-G01). Measured over the real artifacts: of
    # the 486 extraction edges absent from the graph whose endpoints both survive, every one is a
    # `references` edge whose pair the graph draws as calls (480), defines (3), implements (2) or
    # inherits (1) — never as `references` itself. That is the whole of the collapse.
    collapsible = "references"
    strong_relations = frozenset({"calls", "defines", "implements", "inherits"})

    def absorbed(source: str, target: str, relation: str) -> bool:
        """Whether a dropped edge is one the build collapses into a stronger relation."""
        return relation == collapsible and any(
            drawn_relation in strong_relations
            for drawn_relation, _ in drawn_by_pair.get((source, target), set()))

    lost = sorted(
        f"{source} -> {target} as {relation} [{confidence}]"
        for (source, target), variants in relations.items()
        for relation, confidence in variants
        if (relation, confidence) not in drawn_by_pair.get((source, target), set())
        and source in built and target in built
        and source != target
        and not absorbed(source, target, relation))
    if lost:
        raise SystemExit(
            "The extraction holds relationships between nodes the graph kept, but the graph does "
            "not draw them, so it was built from a different extraction:"
            + "\n  " + "\n  ".join(lost[:10]))

    # Above the render ceiling the page draws one node per community, so its ids are community
    # ids; below it they are the graph's own node ids. Either way the sets must match exactly:
    # a subset would let a page drawn from a different partition through.
    drawn_nodes = drawn_page_nodes
    drawn = {str(n.get("id")) for n in drawn_nodes}
    communities = {str(n.get("community")) for n in graph.get("nodes", [])}

    aggregated = drawn == communities
    if not aggregated and drawn != set(built):
        unknown = sorted(drawn - communities - set(built))
        missing = sorted((communities - drawn) if len(drawn) <= len(communities) else (set(built) - drawn))
        raise SystemExit(
            "The page does not draw this graph: it was exported from a different one."
            + (f"\n  drawn but not in the graph: {unknown[:10]}" if unknown else "")
            + (f"\n  in the graph but not drawn: {missing[:10]}" if missing else ""))

    ensure_page_payload_matches(drawn_nodes, html, graph, built, aggregated)


def ensure_page_payload_matches(drawn_nodes: list[dict], html: str, graph: dict,
                                built: dict, aggregated: bool) -> None:
    """Checks what the page says about each node it draws, not only which ones.

    Comparing ids alone left everything else on the page unchecked: a page whose every label read
    <c>FORGED</c> and whose edge array pointed at a node that exists nowhere was accepted as
    derived from this graph (R4-G02).

    <param name="drawn_nodes">The nodes the page draws.</param>
    <param name="html">The page, for its edge array.</param>
    <param name="graph">The graph the page should have come from.</param>
    <param name="built">The graph's nodes, keyed by id.</param>
    <param name="aggregated">Whether the page draws communities rather than nodes.</param>
    <exception cref="SystemExit">Raised when the page's content does not come from this graph.</exception>
    """
    if aggregated:
        expected_labels: dict[str, set[str]] = {}
        for node in graph.get("nodes", []):
            community = str(node.get("community"))
            label = node.get("community_name") or node.get("community_label")
            if label:
                expected_labels.setdefault(community, set()).add(str(label))
    else:
        expected_labels = {node_id: {str(node.get("label"))}
                           for node_id, node in built.items() if node.get("label")}

    mislabelled = []
    for node in drawn_nodes:
        node_id = str(node.get("id"))
        if node_id not in expected_labels:
            continue

        # A node with no label at all was skipped, so removing the label was a way of drawing a
        # node the graph names without saying what it names (R7-G03).
        label = node.get("label")
        if label is None:
            mislabelled.append(f"{node_id}: the page draws it with no label")
        elif str(label) not in expected_labels[node_id]:
            mislabelled.append(f"{node_id}: page says {label!r}")

        if len(mislabelled) >= 10:
            break

    if mislabelled:
        raise SystemExit(
            "The page labels these differently from the graph, so it was exported from a "
            "different one:" + "\n  " + "\n  ".join(mislabelled))

    ensure_page_edges_match(html, graph, aggregated)


def page_edges(html: str) -> list[dict]:
    """Reads the edge array the published page draws.

    <param name="html">The exported page.</param>
    <returns>The drawn edges.</returns>
    <exception cref="SystemExit">Raised when the page carries no readable edge array.</exception>
    """
    return literal_value(html, "RAW_EDGES")


def expected_page_edges(graph: dict, aggregated: bool) -> Counter:
    """Rebuilds, from the graph alone, the edges a page of this graph must draw.

    <param name="graph">The graph the page should have come from.</param>
    <param name="aggregated">Whether the page draws communities rather than nodes.</param>
    <returns>
        A multiset. For a community view: one entry per unordered pair of distinct communities,
        counted once per graph edge crossing it — self-community edges are not drawn. For a node
        view: one entry per graph edge, carrying its direction, relation and confidence.
    </returns>
    """
    if not aggregated:
        return Counter(
            (str(edge.get("source")), str(edge.get("target")),
             str(edge.get("relation") or ""), str(edge.get("confidence") or "EXTRACTED"))
            for edge in graph_edges(graph))

    community_of = {str(n.get("id")): str(n.get("community")) for n in graph.get("nodes", [])}
    counted: Counter = Counter()
    for edge in graph_edges(graph):
        source = community_of.get(str(edge.get("source")))
        target = community_of.get(str(edge.get("target")))
        if source is None or target is None or source == target:
            continue
        counted[tuple(sorted((source, target)))] += 1

    return counted


def drawn_page_edges(edges: list[dict], aggregated: bool) -> Counter:
    """Reads the page's own edges into the same shape <see cref="expected_page_edges" /> produces.

    <param name="edges">The page's edge array.</param>
    <param name="aggregated">Whether the page draws communities rather than nodes.</param>
    <returns>The multiset the page claims.</returns>
    <exception cref="SystemExit">Raised when an aggregated edge does not carry its own count.</exception>
    """
    if not aggregated:
        return Counter(
            (str(edge.get("from")), str(edge.get("to")),
             str(edge.get("label") or ""), str(edge.get("confidence") or "EXTRACTED"))
            for edge in edges)

    counted: Counter = Counter()
    seen: set[tuple[str, str]] = set()
    for edge in edges:
        # The whole label, not a prefix: reading a leading number let the count be split across
        # two entries carrying arbitrary text, with only the sum checked (R7-G03).
        match = re.fullmatch(r"(\d+) cross-community edges", str(edge.get("label") or ""))
        if not match:
            raise SystemExit(
                "An aggregated page edge does not say how many graph edges it stands for, so the "
                f"page cannot be compared with the graph: {edge.get('from')} -> {edge.get('to')} "
                f"labelled {edge.get('label')!r}")

        pair = tuple(sorted((str(edge.get("from")), str(edge.get("to")))))
        if pair in seen:
            raise SystemExit(
                "The page draws two aggregated edges between the same pair of communities, so its "
                f"counts cannot be compared with the graph's: {pair[0]} -- {pair[1]}")

        seen.add(pair)
        counted[pair] += int(match.group(1))

    return counted


def ensure_page_edges_match(html: str, graph: dict, aggregated: bool) -> None:
    """Checks the page draws this graph's edges, and only them.

    <para>
        Checking that both endpoints of every drawn edge were nodes the page drew said nothing
        about direction, relation or how many there were: a page could reverse an edge, invent one
        between two legitimate nodes, or draw no edges at all, and still pass (R4-G02). The page's
        edges are rebuilt from the graph and compared as a multiset instead.
    </para>

    <param name="html">The exported page.</param>
    <param name="graph">The graph the page should have come from.</param>
    <param name="aggregated">Whether the page draws communities rather than nodes.</param>
    <exception cref="SystemExit">Raised when the page's edges are not this graph's edges.</exception>
    """
    expected = expected_page_edges(graph, aggregated)
    drawn = drawn_page_edges(page_edges(html), aggregated)

    if expected == drawn:
        return

    invented = sorted(str(k) for k in (drawn.keys() - expected.keys()))
    missing = sorted(str(k) for k in (expected.keys() - drawn.keys()))
    miscounted = sorted(
        f"{k}: page says {drawn[k]}, graph has {expected[k]}"
        for k in (expected.keys() & drawn.keys()) if expected[k] != drawn[k])

    raise SystemExit(
        "The page does not draw this graph's edges, so it was exported from a different graph or "
        "edited afterwards."
        + (f"\n  drawn but not in the graph: {invented[:10]}" if invented else "")
        + (f"\n  in the graph but not drawn: {missing[:10]}" if missing else "")
        + (f"\n  drawn a different number of times: {miscounted[:10]}" if miscounted else ""))


def provenance_chain(extraction_bytes: bytes, relationship_bytes: bytes, graph_bytes: bytes,
                     html_bytes: bytes, corpus_digest: str, run_id: str) -> dict:
    """Builds a hash chain recording the order the artifacts were derived in.

    Each link folds in the one before it, so the recorded values cannot be reassembled from a
    different combination of artifacts after the fact: the digests in the metadata were three
    independent numbers, and nothing tied them to each other or to the corpus (R4-G02).

    What this does not do is establish where the artifacts came from. The chain travels with the
    metadata that carries it, so anyone who can rewrite one can rewrite both and stay
    self-consistent. It shows that a published map was altered after it was built; it cannot show
    that the build was honest. That needs a signature from a key the publisher does not hold, or a
    rebuild by someone other than the producer.

    <param name="extraction_bytes">The extraction file as published.</param>
    <param name="relationship_bytes">The lossless relationship catalogue derived from it.</param>
    <param name="graph_bytes">The graph file the extraction produced.</param>
    <param name="html_bytes">The page the graph produced.</param>
    <param name="corpus_digest">Digest of the staged corpus.</param>
    <param name="run_id">The staging run the corpus came from.</param>
    <returns>The per-artifact digests and the chained links, as hex.</returns>
    """
    def link(previous: str, payload: bytes) -> str:
        return hashlib.sha256(previous.encode("utf-8") + payload).hexdigest()

    root = hashlib.sha256((corpus_digest + run_id).encode("utf-8")).hexdigest()
    extraction_link = link(root, extraction_bytes)
    relationship_link = link(extraction_link, relationship_bytes)
    graph_link = link(relationship_link, graph_bytes)
    page_link = link(graph_link, html_bytes)

    return {
        "corpus_root": root,
        "extraction_sha256": hashlib.sha256(extraction_bytes).hexdigest(),
        "relationship_sha256": hashlib.sha256(relationship_bytes).hexdigest(),
        "graph_sha256": hashlib.sha256(graph_bytes).hexdigest(),
        "page_sha256": hashlib.sha256(html_bytes).hexdigest(),
        "extraction_link": extraction_link,
        "relationship_link": relationship_link,
        "graph_link": graph_link,
        "page_link": page_link,
    }


def ensure_publish_landed_where_intended(out_dir: Path, expected_key: tuple) -> None:
    """Confirms the published files are in the directory this build created.

    Windows offers no no-follow open, so the path could have been redirected between creating the
    directory and writing into it. That window cannot be closed here (§23.23). What can be done is
    to look afterwards: if the target is no longer the directory that was created, or is now a
    link, the bytes went somewhere this build did not choose and that is said out loud rather than
    reported as a successful publish.

    <param name="out_dir">The publish target.</param>
    <param name="expected_key">The (device, inode) of the directory this build created.</param>
    <exception cref="SystemExit">Raised when the target is not the directory that was created.</exception>
    """
    if out_dir.is_symlink():
        raise SystemExit(
            f"The publish target {out_dir} became a link while the map was being written, so the "
            "files may not be where they were meant to go.")

    try:
        current = out_dir.stat()
    except OSError as error:
        raise SystemExit(
            f"The publish target {out_dir} could not be re-inspected after writing: {error}"
        ) from None

    if (current.st_dev, current.st_ino) != expected_key:
        raise SystemExit(
            f"The publish target {out_dir} is not the directory this build created - it was "
            "replaced while the map was being written, so the files went somewhere else.")

    for name in ("index.html", "relationships.json", "metadata.json"):
        # lstat, not is_file(): `is_file()` follows a link, so it answered about whatever the link
        # pointed at and could not tell a published file from a redirection (R17-G02).
        try:
            leaf = (out_dir / name).lstat()
        except OSError as error:
            raise SystemExit(
                f"The publish target {out_dir} does not hold {name} after writing it, so the "
                f"write did not land where it was directed: {error}") from None

        if not stat.S_ISREG(leaf.st_mode):
            raise SystemExit(
                f"{out_dir / name} is not a regular file after writing it, so the bytes went "
                "somewhere other than the published path.")


def write_regular_file(path: Path, text: str) -> None:
    """Writes text to a path that must be a regular file, never a link to one.

    `Path.write_text` opens an existing name, so a symlink or junction left at `index.html` or
    `metadata.json` would be followed and the bytes written to its target — outside the approved
    directory, whose own identity is bound and re-checked (R17-G02). Refused rather than followed,
    and then written create-new to a sibling and replaced, so the content never travels through a
    name somebody else chose.

    newline="\n" is what makes the recorded digest reproducible: .gitattributes pins these
    artifacts to eol=lf, so bytes written with the platform default would hash differently after a
    clean checkout and verify-public-map.ps1 would fail in CI.
    """
    try:
        existing = path.lstat()
    except FileNotFoundError:
        existing = None
    except OSError as error:
        raise SystemExit(f"{path} could not be inspected before writing: {error}") from None

    if existing is not None and not stat.S_ISREG(existing.st_mode):
        raise SystemExit(
            f"{path} already exists and is not a regular file, so writing it would follow "
            "whatever it points at instead of publishing here. Remove it and rebuild.")

    staging = path.with_name(path.name + ".publishing")
    try:
        # "x" refuses an existing name, so the staging file is one this process created.
        with open(staging, "x", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
        os.replace(staging, path)
    except OSError as error:
        try:
            staging.unlink()
        except OSError:
            pass
        raise SystemExit(f"{path} could not be written: {error}") from None


def resolve_publish_directory(requested: str) -> Path:
    """Canonicalises the publish target and refuses anything outside the documentation tree.

    <para>
        `--out` was joined to the repository root and resolved, which accepts an absolute path and
        `../` alike, and the directory was created immediately — before a single artifact had been
        checked. A failed run therefore left a new directory behind, and a successful one could
        write the page and the metadata anywhere on the machine; the only containment check ran
        after both files were already on disk (R8-G05).
    </para>

    <param name="requested">The caller's --out value.</param>
    <returns>The directory to publish into. It is not created here.</returns>
    """
    docs_root = (REPO_ROOT / "docs").resolve()
    target = Path(requested)
    resolved = (target if target.is_absolute() else REPO_ROOT / target).resolve()

    if resolved != docs_root and docs_root not in resolved.parents:
        raise SystemExit(
            f"Refusing to publish to {resolved}: the map is published inside {docs_root}, and "
            "--out may only name a directory under it.")

    # A link anywhere on the way there points the write somewhere else, whatever the resolved
    # name suggests.
    current = resolved
    while current != docs_root.parent:
        if current.exists() and (current.is_symlink() or current.is_junction()):
            raise SystemExit(
                f"Refusing to publish to {resolved}: {current} is a link, so the files would be "
                "written outside the directory this path names.")
        parent = current.parent
        if parent == current:
            break
        current = parent

    return resolved


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corpus", required=True, help="Directory holding graphify-out/")
    parser.add_argument("--out", default="docs/architecture-map", help="Publish directory")
    parser.add_argument("--spec", required=True,
                        help="Path to the extraction prompt; hashed and recorded as provenance")
    args = parser.parse_args()

    corpus = Path(args.corpus).resolve()
    graph_dir = corpus / "graphify-out"
    if not (graph_dir / ".graphify_extract.json").exists():
        raise SystemExit(f"No extraction found at {graph_dir}/.graphify_extract.json")

    # Nothing from the corpus is executed, and the corpus is verified before anything else is
    # read out of it (R7-G01). `.graphify_python` is the extraction's own note to itself; this
    # script neither trusts it nor runs it.
    pinned = read_pinned_versions()
    versions = preflight(pinned, Path(args.spec).resolve())
    manifest = verify_corpus_manifest(corpus)

    out_dir = resolve_publish_directory(args.out)

    extraction = json.loads((graph_dir / ".graphify_extract.json").read_text(encoding="utf-8"))
    allowed = frozenset(manifest.get("files", {}))
    unapproved = sorted({
        str(item.get("source_file") or "")
        for group in ("nodes", "edges", "hyperedges")
        for item in extraction.get(group, [])
        if not approved(str(item.get("source_file") or ""), corpus, allowed)
    })
    if unapproved:
        raise SystemExit("Extraction contains sources outside the approved corpus:\n  "
                         + "\n  ".join(unapproved[:20]))

    relationships = relationship_catalog(extraction, corpus, manifest["run_id"])
    ensure_relationship_catalog(extraction, relationships, corpus, manifest["run_id"])
    # The catalogue is intentionally complete and can be large. Compact JSON removes only
    # presentation whitespace, keeping every field and duplicate occurrence below the public
    # artifact ceiling without weakening the relationship guarantee.
    relationship_text = json.dumps(
        relationships, ensure_ascii=False, separators=(",", ":")) + "\n"
    relationship_bytes = relationship_text.encode("utf-8")

    graph = json.loads((graph_dir / "graph.json").read_text(encoding="utf-8"))
    html = (graph_dir / "graph.html").read_text(encoding="utf-8")
    ensure_artifacts_share_one_run(extraction, graph, html, corpus)

    # Recorded as a chain rather than three independent digests, so the published metadata says
    # which extraction produced which graph produced which page (R4-G02).
    provenance = provenance_chain(
        (graph_dir / ".graphify_extract.json").read_bytes(),
        relationship_bytes,
        (graph_dir / "graph.json").read_bytes(),
        (graph_dir / "graph.html").read_bytes(),
        manifest["corpus_sha256"], manifest["run_id"])

    # G-06: the page must say what it is a map of. The generated title names a temp path.
    html = re.sub(r"<title>[^<]*</title>",
                  "<title>Aspose MCP Server - Architecture Map</title>", html, count=1)

    # G-08: point the page at the vendored copy of the library. A CDN reference makes the map
    # depend on a third party for availability, leaks every viewer's address to it, and fails
    # under a strict Content Security Policy or offline. The integrity attribute is kept: it
    # documents the exact bytes expected and stops a corrupted or swapped file from running.
    html, replaced = re.subn(
        r'src="https://unpkg\.com/vis-network@[0-9.]+/standalone/umd/vis-network\.min\.js"',
        'src="../assets/vis-network.min.js"', html)
    if not replaced:
        raise SystemExit("Expected a pinned vis-network CDN reference to rewrite; found none.")
    if not (REPO_ROOT / "docs" / "assets" / "vis-network.min.js").exists():
        raise SystemExit("docs/assets/vis-network.min.js is missing; the page would not render.")

    # G-05: neutralise the two places graph text reaches an HTML or script context. This runs
    # before the counts are read so rendered_counts sees the same arrays the page will hold.
    html = harden_page(html)

    commit = source_commit(allowed)
    generated = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    diagnostics = json.loads((graph_dir / ".graphify_diagnostics.json").read_text(encoding="utf-8")) \
        if (graph_dir / ".graphify_diagnostics.json").exists() else {}
    diagnostics = public_diagnostics(diagnostics)
    ensure_extraction_health(diagnostics)

    edges = graph.get("links") or graph.get("edges") or []
    confidence = {}
    for edge in catalog_relationships(relationships):
        key = edge.get("confidence") or "UNKNOWN"
        confidence[key] = confidence.get(key, 0) + 1

    source_nodes, source_edges = len(graph.get("nodes", [])), len(edges)
    source_relationships = relationships["relationship_count"]
    collapsed_relationships = source_relationships - source_edges
    if collapsed_relationships < 0:
        raise SystemExit(
            "The graph contains more edges than the normalized extraction contains relationships.")
    reported_collapsed = diagnostics.get("directed_same_endpoint_collapsed_edges")
    if reported_collapsed is None or int(reported_collapsed) != collapsed_relationships:
        raise SystemExit(
            "The relationship catalogue and Graphify projection disagree about how many "
            "same-endpoint relationships were collapsed: catalogue implies "
            f"{collapsed_relationships}, diagnostics reports {reported_collapsed!r}.")
    drawn_nodes, drawn_edges = rendered_counts(html)
    aggregated = drawn_nodes != source_nodes

    # Saying only "10,928 nodes" over a page that draws 658 community groups reads as a
    # claim about what is on screen (R2-G04), so both numbers are stated and the
    # relationship between them is spelled out.
    scale = (f'analysed {source_nodes:,} nodes / {source_relationships:,} relationships, preserved '
             f'losslessly in relationships.json and projected as {source_edges:,} graph edges; '
             f'drawn here as {drawn_nodes:,} community groups / {drawn_edges:,} links because the '
             'full graph is too large to render'
             if aggregated else
             f'{source_nodes:,} nodes / {source_relationships:,} relationships, preserved '
             f'losslessly in relationships.json and projected as {source_edges:,} graph edges; '
             'drawn in full')

    # The banner used to print the first twelve characters of the commit, which silently
    # dropped the "-dirty" suffix — a map built from an uncommitted tree read exactly like one
    # built from a clean checkout (R3-G04). The suffix is kept and named.
    shown_commit = commit[:12] + ("-dirty (uncommitted working tree)"
                                  if commit.endswith("-dirty") else "")

    banner = (
        '<div style="padding:10px 14px;background:#1f2933;color:#e6edf3;'
        'font:13px/1.5 system-ui,-apple-system,Segoe UI,sans-serif;border-bottom:1px solid #364049">'
        f'<strong>Aspose MCP Server architecture map</strong> &middot; source commit '
        f'<code style="color:#9ecbff">{shown_commit}</code> &middot; generated {generated} &middot; '
        f'{scale} &middot; test code excluded &middot; '
        f'{confidence.get("INFERRED", 0)} of {source_relationships:,} relationships are inferred '
        'rather than '
        'read directly from the source, so treat those as a reading of the code, not a fact about it.'
        '</div>'
    )
    if "<body" in html:
        html = re.sub(r"(<body[^>]*>)", r"\1" + banner, html, count=1)
    else:
        html = banner + html

    # The directory is created here, once every artifact check above has passed, so a
    # refused build leaves nothing behind (R8-G05).
    #
    # Re-resolved first. resolve_publish_directory ran before every artifact was read, checked and
    # hashed, and a local actor able to write under docs/ could turn the target into a link in the
    # meantime — the containment answer would then describe a directory other than the one written
    # to (R9-G02). Asking again costs nothing here and shortens the window to this line.
    if resolve_publish_directory(args.out) != out_dir:
        raise SystemExit(
            f"The publish target {args.out} no longer resolves to {out_dir}. Something changed it "
            "while this build was running; nothing was written.")

    if out_dir.is_symlink() or (out_dir.exists() and out_dir.resolve() != out_dir):
        raise SystemExit(
            f"The publish target {out_dir} is a link, so writing there would put the map somewhere "
            "this check has not approved. Nothing was written.")

    out_dir.mkdir(parents=True, exist_ok=True)

    # The directory now exists. Record what it *is*, so the check after the writes compares an
    # identity rather than a name: on Windows a path can be redirected between the two, and there
    # is no no-follow open to prevent it (§23.23). Preventing it is not available; noticing it is.
    try:
        created_identity = out_dir.stat()
        created_key = (created_identity.st_dev, created_identity.st_ino)
    except OSError as error:
        raise SystemExit(f"The publish target {out_dir} could not be inspected: {error}") from None

    # newline="\n" is what makes the recorded digest reproducible: .gitattributes pins
    # these artifacts to eol=lf, so bytes written with the platform default would hash
    # differently after a clean checkout and verify-public-map.ps1 would fail in CI.
    write_regular_file(out_dir / "index.html", html)
    write_regular_file(out_dir / "relationships.json", relationship_text)

    # The digests above describe graphify-out/graph.html — the page *before* the title change, the
    # CDN rewrite, the hardening and the banner. Nothing recorded what was actually published, so
    # the gate could only compare two of the metadata's own numbers with each other (R7-G03). This
    # is the file a reader downloads, hashed after every transformation.
    published = {
        "index.html": hashlib.sha256((out_dir / "index.html").read_bytes()).hexdigest(),
        "relationships.json": hashlib.sha256(
            (out_dir / "relationships.json").read_bytes()).hexdigest(),
    }
    vendored = REPO_ROOT / "docs" / "assets" / "vis-network.min.js"
    if vendored.exists():
        published["../assets/vis-network.min.js"] = hashlib.sha256(
            vendored.read_bytes()).hexdigest()

    metadata = {
        "source_commit": commit,
        "generated_utc": generated,
        "graphify_package": versions["graphify_package"],
        "graphify_skill_version_observed": versions["graphify_skill_version_observed"],
        "extraction_spec_sha256": versions["extraction_spec_sha256"],
        "corpus_manifest_sha256": manifest["corpus_sha256"],
        "corpus_allowlist_sha256": manifest["allowlist_sha256"],
        "corpus_file_count": manifest["file_count"],
        # Identifies the staging run all three artifacts came from, and chains their digests
        # in derivation order, so a swapped extraction, graph or page is visible after
        # publication too (R3-G03, R4-G02).
        "corpus_run_id": manifest["run_id"],
        "artifact_sha256": {
            ".graphify_extract.json": provenance["extraction_sha256"],
            "relationships.json": provenance["relationship_sha256"],
            "graph.json": provenance["graph_sha256"],
            "graph.html": provenance["page_sha256"],
        },
        "provenance_chain": provenance,
        # What a reader actually receives, hashed after every transformation this script makes.
        "published_sha256": published,
        "corpus_allowlist": {
            "dirs": list(ALLOW_DIRS),
            "files": list(ALLOW_FILES),
            "docs": list(ALLOW_DOCS),
            "excludes_tests": True,
        },
        # Two different graphs, named so neither can be mistaken for the other.
        "source_graph_nodes": source_nodes,
        "source_graph_edges": source_edges,
        "source_relationships": source_relationships,
        "collapsed_relationships": collapsed_relationships,
        "rendered_graph_nodes": drawn_nodes,
        "rendered_graph_edges": drawn_edges,
        "rendered_view": "community-aggregated" if aggregated else "full",
        "communities": len({n.get("community") for n in graph.get("nodes", [])}),
        "edge_confidence": confidence,
        # Unknown means not captured. Zero would claim the extraction was free.
        "semantic_extraction_input_tokens": extraction.get("input_tokens") or None,
        "semantic_extraction_output_tokens": extraction.get("output_tokens") or None,
        "diagnostics": diagnostics,
        "source_files": sorted({
            to_corpus_relative(str(n.get("source_file") or ""), corpus)
            for n in graph.get("nodes", []) if n.get("source_file")
        }),
    }
    write_regular_file(out_dir / "metadata.json",
                       json.dumps(metadata, indent=2, ensure_ascii=False))

    # Every file is written. If the target was replaced between the mkdir above and now, these
    # bytes are somewhere other than where they were meant to be — the window Windows leaves open.
    # It cannot be prevented here, so it is detected: same directory, same files, same digests.
    ensure_publish_landed_where_intended(out_dir, created_key)

    # Record exactly the artifacts this publish consumed, each bound to the bytes it had and to
    # this staging run. Walking the work directory instead laundered whatever happened to be
    # sitting there into the inventory, after which --force would delete it as accounted-for
    # content (R4-G01).
    consumed = {
        f"{STAGING_WORK_DIR}/.graphify_extract.json": "extraction",
        f"{STAGING_WORK_DIR}/graph.json": "graph",
        f"{STAGING_WORK_DIR}/graph.html": "page",
        f"{STAGING_WORK_DIR}/.graphify_diagnostics.json": "diagnostics",
        f"{STAGING_WORK_DIR}/.graphify_python": "interpreter",
    }
    recorded = write_work_inventory(corpus, consumed, str(manifest.get("run_id", "")))
    print(f"Recorded {recorded} consumed artifact(s) in the staging work inventory")

    print(f"Published {out_dir.relative_to(REPO_ROOT)}")
    print(f"  commit {commit}")
    print(f"  analysed {source_nodes} nodes / {source_relationships} relationships, projected as "
          f"{source_edges} edges in "
          f"{metadata['communities']} communities")
    print(f"  rendered {drawn_nodes} nodes / {drawn_edges} edges ({metadata['rendered_view']})")
    print(f"  edge confidence: {confidence}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
