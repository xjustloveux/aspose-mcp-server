<#
.SYNOPSIS
    Refuses to publish anything under docs/ that is not on the manifest.

.DESCRIPTION
    deploy-pages.yml uploads './docs' whole, so every file under it goes live. verify-public-map.ps1
    says exactly that and acts on it — for docs/architecture-map/ only. Nothing applied the same
    reasoning to the rest of the tree, so a file placed anywhere else was published without ever
    being listed or scanned (R13-G01).

    The rules had been written as an xUnit test, which the deploy workflow does not run. A guard
    that is not on the path it guards is a guard in name. This script is that path's guard;
    PublishedDocsInventoryTests exercises the same manifest with negative fixtures.

.PARAMETER DocsDirectory
    The tree about to be published. Defaults to docs.

.PARAMETER ManifestPath
    The manifest naming what may be published. Defaults to graphify/published-docs.json.

.OUTPUTS
    Exit code 0 when every file is accounted for and clean; 1 with a named reason otherwise.
#>
[CmdletBinding()]
param(
    [string]$DocsDirectory = "docs",
    [string]$ManifestPath = (Join-Path $PSScriptRoot "published-docs.json")
)

$ErrorActionPreference = "Stop"

$violations = @()
function Add-Violation([string]$message) { $script:violations += $message }

if (-not (Test-Path $ManifestPath)) {
    Write-Error "No publish manifest at '$ManifestPath'. Refusing to publish an unlisted tree."
    exit 1
}

if (-not (Test-Path $DocsDirectory)) {
    Write-Host "No '$DocsDirectory' directory - nothing to verify."
    exit 0
}

$manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json
$published = @($manifest.published)
$pendingRemoval = @($manifest.pendingRemoval)

if ($published.Count -eq 0) {
    Write-Error "The publish manifest lists no files. An empty manifest would refuse the whole site; fix the manifest rather than the gate."
    exit 1
}

# Repository-relative, posix-separated, so the manifest reads the same on every platform.
# -Force: the workflow uploads the whole tree, hidden files included, and an inventory that
# did not list them let a dotfile go up unaccounted for (R23-G02). Policy: a hidden file is a
# file; it is published only if the manifest names it, and it is scanned like any other.
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$actual = Get-ChildItem -Path $DocsDirectory -Recurse -File -Force |
    ForEach-Object { $_.FullName.Substring($root.Length).TrimStart('\', '/').Replace('\', '/') } |
    Sort-Object

# --- 1. The scan has to have read something ---------------------------------
# Every check below is a comparison against this set; an empty one would make all of them pass.
if ($actual.Count -eq 0) {
    Write-Error "Found no files under '$DocsDirectory'. Refusing rather than reporting a clean tree."
    exit 1
}

# --- 2. Nothing unaccounted for ----------------------------------------------
$accounted = @($published) + @($pendingRemoval | ForEach-Object { $_.path })
foreach ($file in $actual) {
    if ($accounted -notcontains $file) {
        Add-Violation ("'{0}' is under {1} and would be published, and the manifest does not name it. Add it to published if it belongs on the public site, or keep it outside {1}." -f $file, $DocsDirectory)
    }
}

# --- 3. The manifest must describe what is there -----------------------------
foreach ($entry in $published) {
    if ($actual -notcontains $entry) {
        Add-Violation ("'{0}' is listed as published but is not present; the site would link to nothing." -f $entry)
    }
}

foreach ($entry in $pendingRemoval) {
    if ($actual -notcontains $entry.path) {
        Add-Violation ("'{0}' is recorded as pending removal but is already gone. Delete the entry too, or it quietly permits the file's return." -f $entry.path)
    }
}

# --- 4. Division of labour --------------------------------------------------
# This gate is the inventory: nothing goes live that the manifest does not name, and adding a
# file to the manifest is the review point it exists to make unskippable.
#
# Content is the other gate's job. verify-public-map.ps1 reads every file in the published tree
# (R20-G01) and holds each to the rules its kind warrants: the generated map to every rule; a
# hand-written page to credentials, private addresses and identity paths but not to the
# loopback host or system directories it legitimately tells readers about; bytes that are not
# text to credentials and to a drive path only when it continues as one. That "Authored" scope
# is the weaker guarantee, stated: a page may name an absolute path that carries no identity.
#
# What is left unguarded is content a reviewer approved into the manifest that passes those
# rules and is later edited into something that still passes them. Closing that needs review of
# diffs, not a pattern scan.

# --- Result ------------------------------------------------------------------
if ($violations.Count -gt 0) {
    Write-Error ("Refusing to publish '{0}':`n  {1}" -f $DocsDirectory, ($violations -join "`n  "))
    exit 1
}

Write-Host ("Publish manifest gate passed for '{0}' ({1} file(s))." -f $DocsDirectory, $actual.Count)
exit 0
