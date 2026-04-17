#!/usr/bin/env bash
# pack-mcpb.sh — Assemble a per-platform .mcpb bundle for Aspose MCP Server.
#
# Usage:
#   ./deploy/pack-mcpb.sh <platform> <version> <entry_point> <command_expr> <platform_filter>
#
# Arguments:
#   platform        — CI artifact platform name, e.g. windows-x64, linux-x64, macos-arm64, macos-x64
#   version         — Semantic version string, e.g. 1.2.3
#   entry_point     — Manifest server.entry_point value, e.g. server/AsposeMcpServer.exe
#   command_expr    — Manifest mcp_config.command value, e.g. ${__dirname}/server/AsposeMcpServer.exe
#   platform_filter — JSON array string for compatibility.platforms, e.g. ["win32"]
#
# Assumptions:
#   - CWD is repo root when called from CI.
#   - Binary lives at artifacts/<platform>/AsposeMcpServer[.exe].
#   - icon.png has been generated in repo root before this script runs.
#   - The script hard-fails on any bundle-hygiene violation (no continue-on-error).

set -euo pipefail

# ---------------------------------------------------------------------------
# Argument parsing
# ---------------------------------------------------------------------------

if [[ $# -ne 5 ]]; then
    echo "ERROR: Expected 5 arguments, got $#" >&2
    echo "Usage: $0 <platform> <version> <entry_point> <command_expr> <platform_filter>" >&2
    exit 1
fi

PLATFORM="$1"
VERSION="$2"
ENTRY_POINT="$3"
COMMAND_EXPR="$4"
PLATFORM_FILTER="$5"

# Derive binary extension from entry_point: .exe for Windows, empty for Unix.
if [[ "$ENTRY_POINT" == *.exe ]]; then
    BIN_EXT=".exe"
else
    BIN_EXT=""
fi

# Canonical source binary path — always relative to repo root.
BINARY_SOURCE="artifacts/${PLATFORM}/AsposeMcpServer${BIN_EXT}"

STAGING_DIR="/tmp/mcpb-staging-${PLATFORM}"
OUTPUT_FILE="aspose-mcp-server-${PLATFORM}.mcpb"

echo "=== pack-mcpb.sh: ${PLATFORM} v${VERSION} ===" >&2
echo "  Binary source : ${BINARY_SOURCE}" >&2
echo "  Entry point   : ${ENTRY_POINT}" >&2
echo "  Command       : ${COMMAND_EXPR}" >&2
echo "  Platform      : ${PLATFORM_FILTER}" >&2
echo "  Output        : ${OUTPUT_FILE}" >&2

# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------

if [[ ! -f "${BINARY_SOURCE}" ]]; then
    echo "ERROR: Binary not found: ${BINARY_SOURCE}" >&2
    exit 1
fi

if [[ ! -f "icon.png" ]]; then
    echo "ERROR: icon.png not found in repo root. Run the 'Generate icon.png from SVG' step first." >&2
    exit 1
fi

if [[ ! -f "deploy/manifest.template.json" ]]; then
    echo "ERROR: deploy/manifest.template.json not found." >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# Create clean staging directory
# ---------------------------------------------------------------------------

rm -rf "${STAGING_DIR}"
mkdir -p "${STAGING_DIR}/server"

# ---------------------------------------------------------------------------
# Populate staging directory
# ---------------------------------------------------------------------------

# Copy binary into staging/server/
cp "${BINARY_SOURCE}" "${STAGING_DIR}/server/AsposeMcpServer${BIN_EXT}"

# Copy icon.png into staging root
cp "icon.png" "${STAGING_DIR}/icon.png"

# Substitute tokens in manifest template → staging/manifest.json.
# Use sed with a delimiter that does not appear in the substitution values
# (using | as delimiter; COMMAND_EXPR contains / and $). We use a temp file
# approach with Python to avoid shell quoting hazards with ${__dirname}.
python3 - <<PYEOF
import sys

with open('deploy/manifest.template.json', 'r', encoding='utf-8') as fh:
    content = fh.read()

content = content.replace('{{VERSION}}', '${VERSION}')
content = content.replace('{{ENTRY_POINT}}', '${ENTRY_POINT}')
content = content.replace('{{COMMAND}}', '${COMMAND_EXPR}')
content = content.replace('{{PLATFORM_FILTER}}', '${PLATFORM_FILTER}')

with open('${STAGING_DIR}/manifest.json', 'w', encoding='utf-8') as fh:
    fh.write(content)

print('manifest.json written successfully', file=sys.stderr)
PYEOF

# ---------------------------------------------------------------------------
# Gap #4: Ensure executable bit is set on the binary in staging.
# For Unix targets this is required for the bit to be preserved in the ZIP.
# For Windows .exe the chmod is a no-op on Linux staging, which is harmless.
# ---------------------------------------------------------------------------

chmod +x "${STAGING_DIR}/server/AsposeMcpServer${BIN_EXT}"

# ---------------------------------------------------------------------------
# §5.3 Bundle-hygiene guards — all MUST hard-fail on violation
# ---------------------------------------------------------------------------

echo "--- Bundle hygiene checks ---" >&2

# Guard 1: No .lic files
LIC_FILES=$(find "${STAGING_DIR}" -name "*.lic" -type f)
if [[ -n "${LIC_FILES}" ]]; then
    echo "ERROR [Guard 1]: License file(s) found in staging — aborting to prevent license leak:" >&2
    echo "${LIC_FILES}" >&2
    exit 1
fi
echo "  [OK] Guard 1: No .lic files" >&2

# Guard 2: No Aspose XML license content in any file
# Scan for Aspose XML license signature (<License>, <SignedHash>, <Products>).
# We intentionally do NOT print file contents to prevent log exposure.
# Differentiate grep exit codes so genuine I/O errors (exit 2) hard-fail
# rather than being swallowed along with "no match" (exit 1).
set +e
XML_LICENSE_FILES=$(grep -r -l -E '<License>|<SignedHash>|<Products>' "${STAGING_DIR}/")
GREP_STATUS=$?
set -e
if [[ "${GREP_STATUS}" -ge 2 ]]; then
    echo "ERROR [Guard 2]: grep failed (exit ${GREP_STATUS}) — I/O or permission error during license-signature scan" >&2
    exit 1
fi
if [[ -n "${XML_LICENSE_FILES}" ]]; then
    echo "ERROR [Guard 2]: Aspose XML license content detected in staging file(s) — aborting:" >&2
    echo "${XML_LICENSE_FILES}" >&2
    exit 1
fi
echo "  [OK] Guard 2: No Aspose XML license content" >&2

# Guard 3a: No stray image files other than icon.png (svg, ico, icns, bmp, jpg, jpeg, gif, tiff, webp)
STRAY_IMAGES=$(find "${STAGING_DIR}" -type f \( \
    -iname '*.svg' -o \
    -iname '*.ico' -o \
    -iname '*.icns' -o \
    -iname '*.bmp' -o \
    -iname '*.jpg' -o \
    -iname '*.jpeg' -o \
    -iname '*.gif' -o \
    -iname '*.tiff' -o \
    -iname '*.webp' \
\))
if [[ -n "${STRAY_IMAGES}" ]]; then
    echo "ERROR [Guard 3a]: Stray image file(s) found in staging — only icon.png is allowed:" >&2
    echo "${STRAY_IMAGES}" >&2
    exit 1
fi
echo "  [OK] Guard 3a: No stray image files" >&2

# Guard 3b: Exactly one PNG file, and it must be staging/icon.png
PNG_FILES=$(find "${STAGING_DIR}" -type f -iname '*.png')
PNG_COUNT=$(echo "${PNG_FILES}" | grep -c '.' || true)
EXPECTED_PNG="${STAGING_DIR}/icon.png"

if [[ "${PNG_COUNT}" -ne 1 ]]; then
    echo "ERROR [Guard 3b]: Expected exactly 1 PNG file (icon.png), found ${PNG_COUNT}:" >&2
    echo "${PNG_FILES}" >&2
    exit 1
fi

if [[ "${PNG_FILES}" != "${EXPECTED_PNG}" ]]; then
    echo "ERROR [Guard 3b]: PNG file is not at expected location ${EXPECTED_PNG}:" >&2
    echo "${PNG_FILES}" >&2
    exit 1
fi
echo "  [OK] Guard 3b: Exactly one PNG at staging/icon.png" >&2

# Guard 4: Strict file allowlist — staging must contain ONLY these three paths:
#   staging/manifest.json
#   staging/icon.png
#   staging/server/AsposeMcpServer[.exe]
ALLOWED_FILES=(
    "${STAGING_DIR}/manifest.json"
    "${STAGING_DIR}/icon.png"
    "${STAGING_DIR}/server/AsposeMcpServer${BIN_EXT}"
)

ACTUAL_FILES=$(find "${STAGING_DIR}" -type f | sort)
EXPECTED_FILES=$(printf '%s\n' "${ALLOWED_FILES[@]}" | sort)

if [[ "${ACTUAL_FILES}" != "${EXPECTED_FILES}" ]]; then
    echo "ERROR [Guard 4]: Staging directory contains unexpected files." >&2
    echo "  Expected:" >&2
    printf '    %s\n' "${ALLOWED_FILES[@]}" >&2
    echo "  Actual:" >&2
    echo "${ACTUAL_FILES}" | sed 's/^/    /' >&2
    exit 1
fi
echo "  [OK] Guard 4: File allowlist verified" >&2

echo "--- All hygiene checks passed ---" >&2

# ---------------------------------------------------------------------------
# Create .mcpb bundle (plain ZIP archive; executable bit preserved on Linux)
# Capture absolute output path before entering the subshell — "../OUTPUT"
# from /tmp/mcpb-staging-<platform> would resolve to /tmp/, hiding the
# .mcpb from subsequent CI steps that glob from repo root.
# ---------------------------------------------------------------------------

OUTPUT_PATH="$(pwd)/${OUTPUT_FILE}"
(cd "${STAGING_DIR}" && zip -r "${OUTPUT_PATH}" .)

BUNDLE_SIZE=$(du -sh "${OUTPUT_PATH}" | cut -f1)
echo "=== Bundle created: ${OUTPUT_PATH} (${BUNDLE_SIZE}) ===" >&2
