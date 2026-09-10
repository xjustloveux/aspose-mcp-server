"""Make Graphify's public extraction honest before graph construction.

Graphify's C# extractor emits ``imports`` edges for external namespaces but intentionally
does not emit nodes for those third-party/stdlib namespaces. ``build_from_json`` drops those
edges, while the health diagnostic correctly reports their missing endpoints. Publishing the
raw extraction therefore carried thousands of relationships the graph could never represent.

This step removes only that precise, extractor-authored shape and fails closed for every other
dangling edge. The omitted count is saved with the diagnostics, so omission is explicit rather
than mistaken for a complete dependency graph.

Usage:
    python graphify/prepare-public-extraction.py --corpus <staging-dir>
"""
from __future__ import annotations

import argparse
import json
import os
import tempfile
from pathlib import Path

from graphify.diagnostics import diagnose_extraction


EXPECTED_USING_KINDS = frozenset({"namespace", "alias", "static"})


def _is_file_node_label(label: object, source_file: object) -> bool:
    """Match the conservative file-node predicate used by Graphify's builder."""
    if not isinstance(label, str) or not label or not isinstance(source_file, str):
        return False
    normalized = source_file.replace("\\", "/")
    return label == normalized.rsplit("/", 1)[-1] or (
        "/" in label and (normalized == label or normalized.endswith("/" + label)))


def _normalise_semantic_file_mentions(extraction: dict) -> tuple[dict, int]:
    """Apply Graphify's unique AST-file ghost merge before provenance comparison.

    Graphify 0.9.56 removes a non-AST node whose label uniquely names one AST file node,
    rewiring its relationships to that file. Doing this only inside ``build_from_json`` made
    the persisted extraction and graph disagree even though they belonged to one run. This
    mirrors that deliberately narrow rule; ambiguous same-named files remain separate.
    """
    nodes = extraction.get("nodes", [])
    ast_files: list[tuple[str, str]] = []
    for node in nodes:
        if not isinstance(node, dict) or node.get("_origin") != "ast":
            continue
        node_id = node.get("id")
        source_file = node.get("source_file")
        if isinstance(node_id, str) and node_id and _is_file_node_label(
                node.get("label"), source_file):
            ast_files.append((node_id, str(source_file)))

    remap: dict[str, str] = {}
    for node in nodes:
        if not isinstance(node, dict) or node.get("_origin") == "ast":
            continue
        node_id = node.get("id")
        label = node.get("label")
        if not isinstance(node_id, str) or not node_id or not isinstance(label, str):
            continue
        matches = {ast_id for ast_id, ast_source in ast_files
                   if _is_file_node_label(label.strip(), ast_source)}
        if len(matches) == 1 and node_id not in matches:
            remap[node_id] = next(iter(matches))

    if not remap:
        return dict(extraction), 0

    result = dict(extraction)
    result["nodes"] = [node for node in nodes
                       if not isinstance(node, dict) or node.get("id") not in remap]
    result["edges"] = [
        dict(edge,
             source=remap.get(edge.get("source"), edge.get("source")),
             target=remap.get(edge.get("target"), edge.get("target")))
        if isinstance(edge, dict) else edge
        for edge in extraction.get("edges", [])
    ]
    hyperedges: list = []
    for hyperedge in extraction.get("hyperedges", []):
        if not isinstance(hyperedge, dict) or not isinstance(hyperedge.get("nodes"), list):
            hyperedges.append(hyperedge)
            continue
        updated = dict(hyperedge)
        updated["nodes"] = [remap.get(member, member) for member in hyperedge["nodes"]]
        hyperedges.append(updated)
    result["hyperedges"] = hyperedges
    return result, len(remap)


def normalise(extraction: dict) -> tuple[dict, int]:
    """Remove expected external-import edges and refuse every other dangling endpoint.

    Args:
        extraction: Graphify extraction JSON decoded to a dictionary.

    Returns:
        A shallow copy with a representable edge list, and the omitted external-import count.

    Raises:
        SystemExit: If a node has no usable id or an unexplained dangling edge exists.
    """
    extraction, merged_mentions = _normalise_semantic_file_mentions(extraction)
    nodes = extraction.get("nodes", [])
    if not isinstance(nodes, list):
        raise SystemExit("Extraction nodes must be a list.")

    node_ids: set[str] = set()
    for node in nodes:
        if not isinstance(node, dict) or not isinstance(node.get("id"), str) or not node["id"]:
            raise SystemExit("Every extraction node must have a non-empty string id.")
        node_ids.add(node["id"])

    kept: list[dict] = []
    omitted = 0
    unexplained: list[str] = []
    for edge in extraction.get("edges", []):
        if not isinstance(edge, dict):
            unexplained.append(f"non-object edge: {edge!r}")
            continue

        source, target = edge.get("source"), edge.get("target")
        source_exists = source in node_ids
        target_exists = target in node_ids
        if source_exists and target_exists:
            kept.append(edge)
            continue

        metadata = edge.get("metadata") if isinstance(edge.get("metadata"), dict) else {}
        expected_external_import = (
            source_exists
            and not target_exists
            and edge.get("relation") == "imports"
            and edge.get("confidence") == "EXTRACTED"
            and edge.get("_origin") == "ast"
            and metadata.get("using_kind") in EXPECTED_USING_KINDS
            and isinstance(metadata.get("target_fqn"), str)
            and bool(metadata["target_fqn"].strip())
        )
        if expected_external_import:
            omitted += 1
            continue

        unexplained.append(
            f"{source!r} -> {target!r} ({edge.get('relation')!r}, "
            f"{edge.get('confidence')!r}, origin={edge.get('_origin')!r})")

    if unexplained:
        raise SystemExit(
            "Extraction contains dangling edges that are not expected AST external imports:\n  "
            + "\n  ".join(unexplained[:20]))

    result = dict(extraction)
    result["edges"] = kept
    result["_graphify_public_normalization"] = {
        "normalized_semantic_file_mentions": merged_mentions,
    }
    return result, omitted


def _staging_paths(corpus: Path) -> tuple[Path, Path]:
    """Validate a stage marker and return its work and extraction paths."""
    marker_path = corpus / ".graphify-staging"
    manifest_path = corpus / "corpus-manifest.json"
    if not marker_path.is_file() or marker_path.is_symlink() or not manifest_path.is_file():
        raise SystemExit(
            "The corpus is not a Graphify staging directory with a regular marker and manifest.")

    try:
        marker = json.loads(marker_path.read_text(encoding="utf-8"))
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise SystemExit(f"The staging marker or manifest cannot be read: {error}") from None
    if marker.get("schema") != 2 or marker.get("tool") != "graphify/stage-corpus.py":
        raise SystemExit("The staging marker is not one this repository recognises.")
    if marker.get("nonce") != manifest.get("run_id"):
        raise SystemExit("The staging marker and corpus manifest belong to different runs.")

    work = corpus / "graphify-out"
    extraction = work / ".graphify_extract.json"
    if not work.is_dir() or work.is_symlink() or not extraction.is_file() or extraction.is_symlink():
        raise SystemExit("The staged graphify-out or extraction is missing, linked, or not regular.")
    return work, extraction


def _replace_json(path: Path, value: dict) -> None:
    """Atomically replace a JSON artifact inside its existing staging work directory."""
    handle, temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, ensure_ascii=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def main() -> int:
    """Normalize one staged extraction and save its public diagnostic evidence."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--corpus", required=True, help="Graphify staging directory")
    args = parser.parse_args()

    corpus = Path(args.corpus).resolve()
    work, extraction_path = _staging_paths(corpus)
    extraction = json.loads(extraction_path.read_text(encoding="utf-8"))
    cleaned, omitted = normalise(extraction)
    normalization = cleaned.pop("_graphify_public_normalization", {})
    diagnostics = diagnose_extraction(cleaned, directed=True, root=corpus)
    diagnostics["omitted_external_import_edges"] = omitted
    diagnostics["normalized_semantic_file_mentions"] = int(
        normalization.get("normalized_semantic_file_mentions", 0))
    if diagnostics.get("dangling_endpoint_edges", 0):
        raise SystemExit("Normalized extraction still contains dangling endpoint edges.")

    _replace_json(extraction_path, cleaned)
    _replace_json(work / ".graphify_diagnostics.json", diagnostics)
    print(
        f"Prepared public extraction: {len(cleaned.get('nodes', []))} nodes, "
        f"{len(cleaned.get('edges', []))} edges, {omitted} external imports omitted explicitly.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
