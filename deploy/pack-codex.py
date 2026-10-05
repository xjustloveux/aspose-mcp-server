"""Package a Codex marketplace or single plugin using a distribution allowlist.

Run from any directory with Python 3.10+:
    python deploy/pack-codex.py --version 1.2.3 --output publish/marketplace.zip

The archive contains plugin configuration, installation instructions and the
project license. Users supply their platform executable separately through PATH.
Existing output files are preserved, and source manifests are never modified.
"""

import argparse
import json
from pathlib import Path
import re
import sys
import zipfile


PLUGIN_ROOT = "plugins/aspose-mcp-server"
MARKETPLACE_PATH = ".agents/plugins/marketplace.json"
FILES = {
    MARKETPLACE_PATH: MARKETPLACE_PATH,
    f"{PLUGIN_ROOT}/plugin.json": f"{PLUGIN_ROOT}/plugin.json",
    f"{PLUGIN_ROOT}/.codex-plugin/plugin.json": f"{PLUGIN_ROOT}/.codex-plugin/plugin.json",
    f"{PLUGIN_ROOT}/mcp.json": f"{PLUGIN_ROOT}/mcp.json",
    f"{PLUGIN_ROOT}/README.md": f"{PLUGIN_ROOT}/README.md",
    "LICENSE": f"{PLUGIN_ROOT}/LICENSE",
}
VERSION_PATTERN = re.compile(r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)", re.ASCII)


def read_payloads(root: Path, archive_format: str) -> dict[str, bytes]:
    """Read only the regular files allowed for the selected archive format.

    Args:
        root: Resolved repository root containing the distribution sources.
        archive_format: plugin omits the catalog; marketplace includes it.

    Returns:
        Archive entry names mapped to their source bytes.

    Raises:
        ValueError: A source is a symlink or is not a regular file.
        OSError: A source cannot be resolved or read.
    """
    payloads = {}
    for source_name, archive_name in FILES.items():
        if archive_format == "plugin" and source_name == MARKETPLACE_PATH:
            continue
        source = root / source_name
        resolved = source.resolve(strict=True)
        if resolved != source or not resolved.is_file():
            raise ValueError(f"Distribution source must be a regular, non-symlink file: {source_name}")
        payloads[archive_name] = source.read_bytes()
    return payloads


def package(repo_root: Path, output: Path, version: str, archive_format: str = "marketplace") -> None:
    """Validate inputs and write a new ZIP without overwriting files.

    Args:
        repo_root: Root containing the marketplace, plugin and project LICENSE.
        output: New archive path; parent directories are created as needed.
        version: Numeric major.minor.patch version for the packaged manifest.
        archive_format: marketplace (default) preserves the catalog layout;
            plugin puts the plugin files at the archive root and omits the catalog.

    Returns:
        None; writes the new ZIP at output.

    Raises:
        ValueError: The version, source location or JSON configuration is invalid.
        OSError: A source is unavailable or the output already exists.
    """
    if not VERSION_PATTERN.fullmatch(version):
        raise ValueError("Version must be a numeric major.minor.patch release version")
    if archive_format not in ("marketplace", "plugin"):
        raise ValueError("Archive format must be marketplace or plugin")

    root = repo_root.resolve(strict=True)
    payloads = read_payloads(root, archive_format)

    manifest_path = f"{PLUGIN_ROOT}/plugin.json"
    manifest = json.loads(payloads[manifest_path])
    if manifest.get("name") != "aspose-mcp-server":
        raise ValueError("Plugin identity must be aspose-mcp-server")
    manifest["version"] = version
    payloads[manifest_path] = (json.dumps(manifest, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    legacy_path = f"{PLUGIN_ROOT}/.codex-plugin/plugin.json"
    legacy = json.loads(payloads[legacy_path])
    if legacy.get("name") != manifest["name"] or legacy.get("mcpServers") != "./mcp.json":
        raise ValueError("Compatibility manifest must reference the same plugin and MCP configuration")
    legacy["version"] = version
    payloads[legacy_path] = (json.dumps(legacy, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    if archive_format == "marketplace":
        catalog = json.loads(payloads[MARKETPLACE_PATH])
        entries = catalog.get("plugins", [])
        if len(entries) != 1 or entries[0].get("source") != {
            "source": "local", "path": f"./{PLUGIN_ROOT}",
        } or entries[0].get("name") != manifest["name"]:
            raise ValueError("Marketplace must point to the packaged local plugin")
    mcp = json.loads(payloads[f"{PLUGIN_ROOT}/mcp.json"])
    server = mcp.get("mcpServers", {}).get("aspose", {})
    if server.get("type") != "stdio" or server.get("command") != "AsposeMcpServer":
        raise ValueError("MCP configuration must use the local AsposeMcpServer over stdio")

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "x", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, contents in payloads.items():
            if archive_format == "plugin":
                name = name.removeprefix(f"{PLUGIN_ROOT}/")
            archive.writestr(name, contents)


def main() -> int:
    """Parse CLI arguments and package the selected archive format.

    Returns:
        0 on success or 1 for source validation and filesystem failures.
        argparse exits with code 2 for invalid command-line arguments.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--version", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--format", choices=("marketplace", "plugin"), default="marketplace",
                        help="Archive layout (default: marketplace)")
    args = parser.parse_args()
    try:
        package(args.repo_root, args.output, args.version, args.format)
    except (OSError, ValueError) as error:
        print(f"Codex packaging failed: {error}", file=sys.stderr)
        return 1
    print(f"Created Codex {args.format}: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
