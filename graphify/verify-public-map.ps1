<#
.SYNOPSIS
    Publish gate for the public architecture map (G-09).

.DESCRIPTION
    The Pages workflow uploads the whole docs/ directory, so anything copied into
    docs/architecture-map/ goes live. This script is the only thing standing between a
    locally generated graph and the public site: it refuses output that carries a source
    outside the approved corpus, an unreviewed security finding, a local absolute path, a
    secret-shaped string, an internal hostname, or an e-mail address.

    Exit code 0 means the directory is safe to publish. Any other exit code means it is not,
    and the caller must not produce a Pages artifact.

.PARAMETER MapDirectory
    Directory holding the generated map. Defaults to docs/architecture-map.

.PARAMETER MetadataPath
    Path to the generated metadata.json. Defaults to <MapDirectory>/metadata.json.

.PARAMETER RequireCommit
    Commit the published map must describe, for a caller that generates the map in the same job
    and can therefore pin it exactly. Left unset, the map's freshness rests on the corpus digest,
    which is recomputed from the checkout being published and is the stronger check: it compares
    the actual bytes rather than a label.

.PARAMETER MaxFileMegabytes
    Largest single file allowed in the published directory. The full graph belongs in a
    release asset, not in Git history (G-05).
#>
param(
    [string]$MapDirectory = "docs/architecture-map",
    [string]$MetadataPath = "",
    [string]$RequireCommit = "",
    [double]$MaxFileMegabytes = 10
)

$ErrorActionPreference = "Stop"
$violations = [System.Collections.Generic.List[string]]::new()

function Add-Violation([string]$message) {
    $script:violations.Add($message)
}

if (-not (Test-Path $MapDirectory)) {
    # A missing map matters only when the site advertises it. Silently succeeding in that
    # case would publish navigation and a sitemap entry pointing at a page that is not there.
    $docsRoot = Split-Path $MapDirectory -Parent
    $referrers = @()
    if (Test-Path $docsRoot) {
        $referrers = Get-ChildItem -Path $docsRoot -File -Recurse -Force -Include *.html, *.xml |
            Where-Object { (Get-Content $_.FullName -Raw) -match "architecture-map" } |
            ForEach-Object { $_.Name }
    }
    if ($referrers.Count -gt 0) {
        Write-Error ((("No map directory at '{0}', but {1} published file(s) link to it ({2}). " +
                       "Publishing would leave those links pointing at a page that does not exist. " +
                       "Build the map with graphify/build-public-map.py, or remove the references.")) -f
                      $MapDirectory, $referrers.Count, ($referrers -join ", "))
        exit 1
    }
    Write-Host "No map directory at '$MapDirectory' and nothing links to one - nothing to verify."
    exit 0
}

if ([string]::IsNullOrEmpty($MetadataPath)) {
    $MetadataPath = Join-Path $MapDirectory "metadata.json"
}

# --- 1. Required files -------------------------------------------------------
$indexPath = Join-Path $MapDirectory "index.html"
$relationshipsPath = Join-Path $MapDirectory "relationships.json"
if (-not (Test-Path $indexPath)) { Add-Violation "index.html is missing from $MapDirectory" }
if (-not (Test-Path $relationshipsPath)) { Add-Violation "relationships.json is missing from $MapDirectory" }
if (-not (Test-Path $MetadataPath)) { Add-Violation "metadata.json is missing from $MapDirectory" }

# --- 2. Size ceiling ---------------------------------------------------------
foreach ($file in Get-ChildItem -Path $MapDirectory -Recurse -File -Force) {
    $mb = $file.Length / 1MB
    if ($mb -gt $MaxFileMegabytes) {
        Add-Violation ("{0} is {1:N2} MB, above the {2} MB publish ceiling" -f $file.Name, $mb, $MaxFileMegabytes)
    }
}

# --- 2b. Vendored assets the page depends on --------------------------------
if (Test-Path $indexPath) {
    $indexHtml = Get-Content $indexPath -Raw
    foreach ($match in [regex]::Matches($indexHtml, 'src="(\.\./assets/[^"]+)"')) {
        $assetRelative = $match.Groups[1].Value -replace '\.\./', ''
        $assetPath = Join-Path (Split-Path $MapDirectory -Parent) $assetRelative
        if (-not (Test-Path $assetPath)) {
            Add-Violation "index.html references a vendored asset that is not present: $assetRelative"
            continue
        }
        # An integrity attribute that no longer matches the file makes the browser refuse it,
        # so a mismatch is a broken page, not a warning.
        $integrity = [regex]::Match($indexHtml, 'integrity="sha384-([^"]+)"')
        if ($integrity.Success) {
            $sha = [System.Security.Cryptography.SHA384]::Create()
            $actual = [Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($assetPath)))
            if ($actual -ne $integrity.Groups[1].Value) {
                Add-Violation "$assetRelative does not match the integrity hash declared in index.html"
            }
        }
    }
}

# --- 3. Metadata source allowlist -------------------------------------------
# The canonical corpus membership comes from corpus_allowlist.py by way of corpus-digest.py.
# This script used to keep its own prefix list, which accepted anything under docs/ - the
# map's own output included - while the canonical allowlist named nine pages (R2-G03).
if (Test-Path $MetadataPath) {
    $metadata = Get-Content $MetadataPath -Raw | ConvertFrom-Json

    # Every one of these is measured at generation time. Their absence, or a placeholder
    # value, means the map cannot be traced to the toolchain and sources that produced it.
    $requiredFields = @(
        "source_commit", "generated_utc", "graphify_package",
        "graphify_skill_version_observed", "extraction_spec_sha256",
        "corpus_manifest_sha256", "corpus_allowlist_sha256", "corpus_file_count",
        "corpus_run_id",
        "source_graph_nodes", "source_graph_edges", "source_relationships",
        "collapsed_relationships",
        "rendered_graph_nodes", "rendered_graph_edges",
        "source_files", "diagnostics"
    )
    foreach ($required in $requiredFields) {
        if (-not $metadata.PSObject.Properties.Name.Contains($required)) {
            Add-Violation "metadata.json is missing required field '$required'"
            continue
        }
        $value = $metadata.$required
        if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value) -or
            [string]$value -eq "unknown") {
            Add-Violation "metadata.json field '$required' is empty or unknown; provenance must be observed, not assumed"
        }
    }

    # Graphify intentionally does not create nodes for external C# namespaces. The preparation
    # step removes only those precise AST import edges and records how many were omitted; every
    # edge left for the public graph must have real endpoints. Requiring the diagnostic here keeps
    # a skipped preparation step from silently publishing a graph that discarded relationships.
    $diagnostics = $metadata.diagnostics
    if ($null -ne $diagnostics) {
        foreach ($required in @("dangling_endpoint_edges", "missing_endpoint_edges",
                                "self_loop_edges", "directed_same_endpoint_collapsed_edges",
                                "omitted_external_import_edges",
                                "normalized_semantic_file_mentions")) {
            if (-not $diagnostics.PSObject.Properties.Name.Contains($required)) {
                Add-Violation "metadata.json diagnostics is missing '$required'"
            }
        }
        foreach ($mustBeZero in @("dangling_endpoint_edges", "missing_endpoint_edges",
                                  "self_loop_edges")) {
            if ($diagnostics.PSObject.Properties.Name.Contains($mustBeZero) -and
                [int]$diagnostics.$mustBeZero -ne 0) {
                Add-Violation ("metadata.json diagnostics reports {0}={1}; public edges must " +
                               "all be representable" -f $mustBeZero,$diagnostics.$mustBeZero)
            }
        }
        if ($diagnostics.PSObject.Properties.Name.Contains("omitted_external_import_edges") -and
            [int]$diagnostics.omitted_external_import_edges -lt 0) {
            Add-Violation "metadata.json diagnostics has a negative omitted external-import count"
        }
        if ($diagnostics.PSObject.Properties.Name.Contains("normalized_semantic_file_mentions") -and
            [int]$diagnostics.normalized_semantic_file_mentions -lt 0) {
            Add-Violation "metadata.json diagnostics has a negative semantic file-mention count"
        }
    }

    # The run id says which staging run produced the corpus; these digests say which
    # extraction, graph and page were built from it. Without them a later swap of any one
    # artifact leaves nothing in the published output to contradict it (R3-G03).
    $artifacts = $metadata.artifact_sha256
    if ($null -eq $artifacts) {
        Add-Violation "metadata.json records no artifact_sha256; the published page cannot be tied to the extraction it came from"
    }
    else {
        foreach ($artifact in @('.graphify_extract.json', 'relationships.json',
                                'graph.json', 'graph.html')) {
            $digest = [string]$artifacts.$artifact
            if ($digest -notmatch '^[0-9a-f]{64}$') {
                Add-Violation "metadata.json artifact_sha256 has no valid digest for '$artifact'"
            }
        }
    }

    # Three independent digests said nothing about which artifact produced which. The chain folds
    # each into the next, so the recorded values cannot be reassembled from a different
    # combination after the fact (R4-G02).
    $chain = $metadata.provenance_chain
    if ($null -eq $chain) {
        Add-Violation "metadata.json records no provenance_chain; the artifacts are not tied to each other"
    }
    else {
        foreach ($link in @('corpus_root', 'extraction_sha256', 'relationship_sha256',
                            'graph_sha256', 'page_sha256', 'extraction_link',
                            'relationship_link', 'graph_link', 'page_link')) {
            if ([string]$chain.$link -notmatch '^[0-9a-f]{64}$') {
                Add-Violation "metadata.json provenance_chain has no valid '$link'"
            }
        }

        # The two records of the same bytes must agree with each other.
        foreach ($pair in @(
            @{ Artifact = '.graphify_extract.json'; Link = 'extraction_sha256' },
            @{ Artifact = 'relationships.json'; Link = 'relationship_sha256' },
            @{ Artifact = 'graph.json'; Link = 'graph_sha256' },
            @{ Artifact = 'graph.html'; Link = 'page_sha256' })) {
            if ($null -ne $artifacts -and
                [string]$artifacts.($pair.Artifact) -ne [string]$chain.($pair.Link)) {
                Add-Violation ("metadata.json artifact_sha256['{0}'] disagrees with provenance_chain.{1}" -f
                               $pair.Artifact, $pair.Link)
            }
        }

        if ([string]$chain.corpus_root -eq [string]$chain.extraction_link) {
            Add-Violation "metadata.json provenance_chain does not fold the extraction into the corpus root"
        }
    }

    # Everything above compares the metadata with itself. The page a reader downloads is the one
    # this script must hash: metadata recorded graphify-out/graph.html, and the published
    # index.html is that file after a title change, a CDN rewrite, hardening and a banner, so a
    # byte edited afterwards contradicted nothing (R7-G03).
    $publishedDigests = $metadata.published_sha256
    if ($null -eq $publishedDigests) {
        Add-Violation "metadata.json records no published_sha256; the file on disk is tied to nothing"
    }
    else {
        foreach ($entry in $publishedDigests.PSObject.Properties) {
            $target = Join-Path $MapDirectory $entry.Name
            if (-not (Test-Path -LiteralPath $target)) {
                Add-Violation ("metadata.json records a digest for '{0}', which is not published" -f $entry.Name)
                continue
            }

            $actual = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($actual -ne [string]$entry.Value) {
                Add-Violation ("'{0}' does not match the digest recorded when it was published" -f $entry.Name)
            }
        }

        if (-not $publishedDigests.PSObject.Properties.Name.Contains('index.html')) {
            Add-Violation "metadata.json published_sha256 does not cover index.html"
        }
        if (-not $publishedDigests.PSObject.Properties.Name.Contains('relationships.json')) {
            Add-Violation "metadata.json published_sha256 does not cover relationships.json"
        }
    }

    # The clustering graph is intentionally a simple DiGraph. This sidecar is the authoritative
    # lossless relationship set, so every entry and every collapse count is checked independently
    # of the rendered projection (R31-G01).
    $relationshipCatalog = $null
    if (Test-Path $relationshipsPath) {
        try { $relationshipCatalog = Get-Content $relationshipsPath -Raw | ConvertFrom-Json }
        catch { Add-Violation "relationships.json is not valid JSON" }
    }
    if ($null -ne $relationshipCatalog) {
        if ([int]$relationshipCatalog.schema -ne 1 -or
            [string]$relationshipCatalog.format -ne 'grouped-fields') {
            Add-Violation "relationships.json has an unsupported schema or storage format"
        }
        if ([string]$relationshipCatalog.corpus_run_id -ne [string]$metadata.corpus_run_id) {
            Add-Violation "relationships.json belongs to a different staging run than metadata.json"
        }
        $orderedPairs = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::Ordinal)
        $relationshipSources = [System.Collections.Generic.List[string]]::new()
        $relationshipCount = 0
        foreach ($group in @($relationshipCatalog.groups)) {
            $fields = [string[]]@($group.fields)
            $fieldSet = [System.Collections.Generic.HashSet[string]]::new(
                $fields, [System.StringComparer]::Ordinal)
            $sourceIndex = [Array]::IndexOf($fields, 'source')
            $targetIndex = [Array]::IndexOf($fields, 'target')
            $relationIndex = [Array]::IndexOf($fields, 'relation')
            $sourceFileIndex = [Array]::IndexOf($fields, 'source_file')
            if ($fields.Count -eq 0 -or $fieldSet.Count -ne $fields.Count -or
                $sourceIndex -lt 0 -or $targetIndex -lt 0 -or $relationIndex -lt 0) {
                Add-Violation "relationships.json contains a group with invalid or incomplete fields"
                continue
            }
            foreach ($rowValue in @($group.rows)) {
                $row = @($rowValue)
                $relationshipCount++
                if ($row.Count -ne $fields.Count) {
                    Add-Violation "relationships.json contains a row that does not match its fields"
                    continue
                }
                $source = [string]$row[$sourceIndex]
                $target = [string]$row[$targetIndex]
                $relation = [string]$row[$relationIndex]
                if ([string]::IsNullOrWhiteSpace($source) -or
                    [string]::IsNullOrWhiteSpace($target) -or
                    [string]::IsNullOrWhiteSpace($relation)) {
                    Add-Violation "relationships.json contains a relationship without source, target, or relation"
                    continue
                }
                $null = $orderedPairs.Add(("{0}:{1}{2}" -f $source.Length, $source, $target))
                if ($sourceFileIndex -ge 0) {
                    $relationshipSources.Add([string]$row[$sourceFileIndex])
                }
            }
        }
        if ([int]$relationshipCatalog.relationship_count -ne $relationshipCount) {
            Add-Violation "relationships.json relationship_count does not match its rows"
        }
        if ([int]$metadata.source_relationships -ne $relationshipCount) {
            Add-Violation "metadata.json source_relationships does not match relationships.json"
        }
        $collapsed = $relationshipCount - $orderedPairs.Count
        if ([int]$metadata.source_graph_edges -ne $orderedPairs.Count) {
            Add-Violation "metadata.json source_graph_edges is not the simple-graph projection of relationships.json"
        }
        if ([int]$metadata.collapsed_relationships -ne $collapsed) {
            Add-Violation "metadata.json collapsed_relationships does not match relationships.json"
        }
        if ($null -ne $diagnostics -and
            $diagnostics.PSObject.Properties.Name.Contains('directed_same_endpoint_collapsed_edges') -and
            [int]$diagnostics.directed_same_endpoint_collapsed_edges -ne $collapsed) {
            Add-Violation "metadata.json diagnostics disagrees with the relationship collapse count"
        }
    }

    # The map is built from a staged corpus; the manifest digest recorded in the metadata
    # must still describe the tree being published. Every failure below is a violation, not
    # a note: this comparison used to be skipped whenever the helper could not run, so a
    # broken interpreter on CI turned the freshness check off and still exited 0 (R2-G02).
    $canonical = $null
    $digestOutput = & python (Join-Path $PSScriptRoot "corpus-digest.py") 2>&1
    if ($LASTEXITCODE -ne 0) {
        Add-Violation "corpus-digest.py exited with code $LASTEXITCODE; corpus provenance cannot be verified"
    }
    elseif ([string]::IsNullOrWhiteSpace(($digestOutput -join ""))) {
        Add-Violation "corpus-digest.py produced no output; corpus provenance cannot be verified"
    }
    else {
        try {
            $canonical = ($digestOutput -join "`n") | ConvertFrom-Json
        }
        catch {
            Add-Violation "corpus-digest.py output could not be parsed as JSON; corpus provenance cannot be verified"
        }
    }

    if ($null -ne $canonical -and [string]::IsNullOrWhiteSpace([string]$canonical.corpus_sha256)) {
        Add-Violation "corpus-digest.py output carries no corpus_sha256; corpus provenance cannot be verified"
    }
    elseif ($null -ne $canonical -and $canonical.corpus_sha256 -ne [string]$metadata.corpus_manifest_sha256) {
        Add-Violation ("The published map was built from a different corpus than the " +
                       "current tree (metadata {0}, tree {1}). Re-stage, re-extract and rebuild." -f
                       ([string]$metadata.corpus_manifest_sha256).Substring(0, 12),
                       ([string]$canonical.corpus_sha256).Substring(0, 12))
    }

    $sources = @($metadata.source_files | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($sources.Count -eq 0) {
        Add-Violation "metadata.json lists no source files; a map with no recorded sources cannot be reviewed"
    }
    elseif ($null -ne $canonical -and $null -ne $canonical.files) {
        $allowed = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]@($canonical.files), [System.StringComparer]::OrdinalIgnoreCase)
        foreach ($source in $sources) {
            $normalized = ([string]$source) -replace '\\', '/'
            if (-not $allowed.Contains($normalized)) {
                Add-Violation "metadata.json lists a source outside the approved corpus: $normalized"
            }
        }
        if ($null -ne $relationshipCatalog) {
            foreach ($sourceValue in $relationshipSources) {
                $relationshipSource = ([string]$sourceValue) -replace '\\', '/'
                if (-not [string]::IsNullOrWhiteSpace($relationshipSource) -and
                    -not $allowed.Contains($relationshipSource)) {
                    Add-Violation "relationships.json lists a source outside the approved corpus: $relationshipSource"
                }
            }
        }
    }
}

# A page literal, found the way the generator finds it: `const NAME = <array or object>;`,
# scanned with knowledge of JSON strings and balanced brackets. `\[.*?\];` ended at the first
# `];` a label happened to contain, and then this gate checked the same truncated prefix the
# hardener had escaped (R22-G01). Returns the literal text, or $null with the reason in
# $script:LiteralError.
function Get-ScriptLiteral([string]$Html, [string]$Name) {
    $script:LiteralError = $null
    # A declaration starts a line, and there is exactly one: the first `const NAME =` anywhere
    # could be text inside an earlier literal's string, and checking that decoy passed a page
    # whose real literal was raw (R23-G01).
    $heads = [regex]::Matches($Html, ('(?m)^[ \t]*const\s+{0}\s*=\s*' -f [regex]::Escape($Name)))
    if ($heads.Count -eq 0) { $script:LiteralError = "carries no $Name declaration at the start of a line"; return $null }
    if ($heads.Count -gt 1) { $script:LiteralError = "declares $Name $($heads.Count) times"; return $null }
    $head = $heads[0]
    $script:LiteralStart = $head.Index + $head.Length
    $start = $head.Index + $head.Length
    if ($start -ge $Html.Length -or ($Html[$start] -ne '[' -and $Html[$start] -ne '{')) {
        $script:LiteralError = "$Name does not begin with an array or an object"; return $null
    }
    $stack = [System.Collections.Generic.Stack[char]]::new()
    $inString = $false
    $escaped = $false
    for ($i = $start; $i -lt $Html.Length; $i++) {
        $c = $Html[$i]
        if ($inString) {
            if ($escaped) { $escaped = $false }
            elseif ($c -eq '\') { $escaped = $true }
            elseif ($c -eq '"') { $inString = $false }
            continue
        }
        if ($c -eq '"') { $inString = $true }
        elseif ($c -eq '[') { $stack.Push(']') }
        elseif ($c -eq '{') { $stack.Push('}') }
        elseif ($c -eq ']' -or $c -eq '}') {
            if ($stack.Count -eq 0 -or $stack.Pop() -ne $c) {
                $script:LiteralError = "$Name closes a bracket it never opened"; return $null
            }
            if ($stack.Count -eq 0) {
                $end = $i + 1
                if (-not [regex]::IsMatch($Html.Substring($end), '^\s*;')) {
                    $script:LiteralError = "$Name is not terminated by a semicolon"; return $null
                }
                return $Html.Substring($start, $end - $start)
            }
        }
    }
    $script:LiteralError = "$Name never closes"
    return $null
}

# --- 3b. Rendered counts ----------------------------------------------------
# The page draws a community-aggregated view once the graph passes graphify's render
# ceiling, so the totals in metadata.json describe a different graph than the one on
# screen. Both are published; this checks the rendered pair against the actual arrays so
# the page can never claim to draw more than it does (R2-G04).
if ((Test-Path $indexPath) -and (Test-Path $MetadataPath) -and $null -ne $metadata) {
    foreach ($pair in @(
        @{ Name = 'RAW_NODES'; Field = 'rendered_graph_nodes' },
        @{ Name = 'RAW_EDGES'; Field = 'rendered_graph_edges' })) {
        $literal = Get-ScriptLiteral $indexHtml $pair.Name
        if ($null -eq $literal) {
            Add-Violation "index.html $($script:LiteralError); the rendered size cannot be verified"
            continue
        }
        try {
            $actual = @($literal | ConvertFrom-Json).Count
        }
        catch {
            Add-Violation "index.html $($pair.Name) could not be parsed; the rendered size cannot be verified"
            continue
        }
        $claimed = [int]$metadata.($pair.Field)
        if ($actual -ne $claimed) {
            Add-Violation ("metadata.json says {0}={1} but index.html renders {2}" -f $pair.Field, $claimed, $actual)
        }
    }
}

# --- 3c. Script-context safety ----------------------------------------------
# Node labels and community names come from source symbols and documentation headings, so
# markup in them is not under the generator's control. Inside a <script> element an angle
# bracket can close the element early, so the data arrays must carry none (R2-G05).
if (Test-Path $indexPath) {
    # The same four literals the generator hardens; a name here and not there is drift
    # (R21-G04; `hyperedges` since R21-G03's guard found it unhardened). The whole literal,
    # by the same scan the generator uses, must parse as JSON and carry none of the characters
    # that end a script element or a JavaScript line (R22-G01). And no literal may lie inside
    # another: a declaration found inside a string is a decoy (R23-G01).
    $ranges = @{}
    foreach ($name in @('RAW_NODES', 'RAW_EDGES', 'LEGEND', 'hyperedges')) {
        $literal = Get-ScriptLiteral $indexHtml $name
        if ($null -ne $literal) { $ranges[$name] = @($script:LiteralStart, ($script:LiteralStart + $literal.Length)) }
    }
    $ordered = @($ranges.GetEnumerator() | Sort-Object { $_.Value[0] })
    for ($k = 1; $k -lt $ordered.Count; $k++) {
        if ($ordered[$k].Value[0] -lt $ordered[$k - 1].Value[1]) {
            Add-Violation "index.html $($ordered[$k].Key) is declared inside the $($ordered[$k - 1].Key) literal; a declaration inside a string is a decoy"
        }
    }
    foreach ($name in @('RAW_NODES', 'RAW_EDGES', 'LEGEND', 'hyperedges')) {
        $literal = Get-ScriptLiteral $indexHtml $name
        if ($null -eq $literal) {
            Add-Violation "index.html $($script:LiteralError); its script context cannot be verified"
            continue
        }
        try { $null = $literal | ConvertFrom-Json -AsHashtable }
        catch { Add-Violation "index.html $name is not valid JSON; its script context cannot be verified"; continue }
        foreach ($forbidden in @(
            @{ Char = '<'; Why = 'a raw < inside a script element' },
            @{ Char = '>'; Why = 'a raw > inside a script element' },
            @{ Char = [string][char]0x2028; Why = 'a raw U+2028 line separator' },
            @{ Char = [string][char]0x2029; Why = 'a raw U+2029 paragraph separator' })) {
            if ($literal.Contains($forbidden.Char)) {
                Add-Violation "index.html $name contains $($forbidden.Why); graph text must be unicode-escaped"
            }
        }
    }
}

# --- 3d. Metadata must agree with the page it describes ---------------------
# The rendered pair is checked against the arrays above, but the source pair, the view name
# and the banner were four independent copies of the same facts: editing any one of them
# left the other three agreeing with each other and the gate silent (R3-G05). Every number
# the banner states is recomputed here from the metadata and the arrays it claims to
# summarise, so no single value can be changed on its own.
function Get-BannerNumbers([string]$text, [string]$pattern) {
    $match = [regex]::Match($text, $pattern)
    if (-not $match.Success) { return $null }
    return @($match.Groups | Select-Object -Skip 1 | ForEach-Object { [int]($_.Value -replace ',', '') })
}

$banner = $null
if ((Test-Path $indexPath) -and $null -ne $metadata) {
    $bannerMatch = [regex]::Match($indexHtml,
        '<div style="padding:10px 14px;background:#1f2933.*?</div>',
        [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $bannerMatch.Success) {
        Add-Violation "index.html carries no provenance banner; its numbers cannot be cross-checked"
    }
    else {
        $banner = $bannerMatch.Value

        # The template lays <body> out as a flex row of #graph and #sidebar. A banner placed in
        # that row took the whole width and left the graph zero pixels wide, so the page listed
        # every node and drew none of them, and every check above still passed.
        $layoutRule = '<style>body{display:grid;grid-template-columns:minmax(0,1fr) 280px;' +
                      'grid-template-rows:auto minmax(0,1fr)}#graph,#sidebar{min-width:0;min-height:0}</style></head>'
        if ($banner -notmatch [regex]::Escape('grid-column:1/-1') -or
            -not $indexHtml.Contains($layoutRule)) {
            Add-Violation "index.html banner shares the graph's row, so the graph is drawn zero pixels wide"
        }
    }
}

if ($null -ne $banner) {
    $sourceNodes = [int]$metadata.source_graph_nodes
    $sourceEdges = [int]$metadata.source_graph_edges
    $sourceRelationships = [int]$metadata.source_relationships
    $renderedNodes = [int]$metadata.rendered_graph_nodes
    $renderedEdges = [int]$metadata.rendered_graph_edges

    if ($renderedNodes -gt $sourceNodes -or $renderedEdges -gt $sourceEdges) {
        Add-Violation ("metadata.json renders more than it analysed ({0}/{1} rendered from {2}/{3} source)" -f
                       $renderedNodes, $renderedEdges, $sourceNodes, $sourceEdges)
    }

    # The view name is derived, not declared: aggregation is exactly the case where the page
    # draws a different number of nodes than the graph holds.
    $expectedView = if ($renderedNodes -ne $sourceNodes) { "community-aggregated" } else { "full" }
    if ([string]$metadata.rendered_view -ne $expectedView) {
        Add-Violation ("metadata.json says rendered_view='{0}' but the counts describe '{1}'" -f
                       $metadata.rendered_view, $expectedView)
    }

    $aggregated = $expectedView -eq "community-aggregated"
    if ($aggregated -and $banner -notmatch 'drawn here as') {
        Add-Violation "index.html draws a community-aggregated view but its banner does not say so"
    }
    if (-not $aggregated -and $banner -notmatch 'drawn in full') {
        Add-Violation "index.html draws the full graph but its banner does not say so"
    }

    if ($banner -notmatch 'preserved losslessly in relationships\.json') {
        Add-Violation "index.html banner does not disclose the lossless relationship artifact"
    }

    $stated = if ($aggregated) {
        Get-BannerNumbers $banner ('analysed ([\d,]+) nodes / ([\d,]+) relationships, preserved ' +
                                   'losslessly in relationships\.json and projected as ([\d,]+) graph edges; ' +
                                   'drawn here as ([\d,]+) community groups / ([\d,]+) links')
    }
    else {
        Get-BannerNumbers $banner ('([\d,]+) nodes / ([\d,]+) relationships, preserved ' +
                                   'losslessly in relationships\.json and projected as ([\d,]+) graph edges; ' +
                                   'drawn in full')
    }

    if ($null -eq $stated) {
        Add-Violation "index.html banner does not state the graph size in the expected form"
    }
    else {
        $expected = if ($aggregated) {
            @($sourceNodes, $sourceRelationships, $sourceEdges, $renderedNodes, $renderedEdges)
        }
        else { @($sourceNodes, $sourceRelationships, $sourceEdges) }
        for ($i = 0; $i -lt $expected.Count; $i++) {
            if ($stated[$i] -ne $expected[$i]) {
                Add-Violation ("index.html banner states {0} where metadata.json says {1}" -f
                               $stated[$i], $expected[$i])
            }
        }
    }

    # The inference disclaimer is the one claim a reader is most likely to act on, so it is
    # checked against the recorded confidence breakdown rather than taken on trust.
    $inferredStated = Get-BannerNumbers $banner '([\d,]+) of ([\d,]+) relationships are inferred'
    $inferredRecorded = 0
    $confidenceTotal = 0
    if ($null -ne $metadata.edge_confidence) {
        foreach ($property in $metadata.edge_confidence.PSObject.Properties) {
            $confidenceTotal += [int]$property.Value
            if ($property.Name -eq 'INFERRED') { $inferredRecorded = [int]$property.Value }
        }
    }
    if ($confidenceTotal -ne $sourceRelationships) {
        Add-Violation ("metadata.json edge_confidence totals {0} but source_relationships is {1}" -f
                       $confidenceTotal, $sourceRelationships)
    }
    if ($null -eq $inferredStated) {
        Add-Violation "index.html banner does not state how many relationships are inferred"
    }
    elseif ($inferredStated[0] -ne $inferredRecorded -or
            $inferredStated[1] -ne $sourceRelationships) {
        Add-Violation ("index.html banner claims {0} of {1} inferred; metadata.json records {2} of {3}" -f
                       $inferredStated[0], $inferredStated[1], $inferredRecorded, $sourceRelationships)
    }

    # ConvertFrom-Json turns an ISO-8601 string into a [datetime], and stringifying that gives
    # the host's local format rather than the text the page carries, so it is written back out
    # in the form the generator used.
    $generated = $metadata.generated_utc
    $generatedText = if ($generated -is [datetime]) {
        $generated.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    else {
        [string]$generated
    }

    if ($banner -notmatch [regex]::Escape($generatedText)) {
        Add-Violation "index.html banner does not carry the generation time recorded in metadata.json"
    }
}

# --- 3e. The commit the map claims ------------------------------------------
# The recorded commit was only required to be non-empty, so a map built weeks ago, or from
# an uncommitted tree, published as a description of today's source. The banner showed the
# first twelve characters, which hid the "-dirty" suffix entirely (R3-G04).
if ($null -ne $metadata) {
    $claimedCommit = [string]$metadata.source_commit
    $isDirty = $claimedCommit.EndsWith("-dirty")

    if ($null -ne $banner) {
        $shortCommit = $claimedCommit -replace '-dirty$', ''
        if ($shortCommit.Length -ge 12) { $shortCommit = $shortCommit.Substring(0, 12) }
        if ($banner -notmatch [regex]::Escape($shortCommit)) {
            Add-Violation "index.html banner does not show the commit recorded in metadata.json"
        }
        if ($isDirty -and $banner -notmatch 'dirty') {
            Add-Violation "metadata.json records an uncommitted tree but the banner does not disclose it"
        }
    }

    # Freshness is established by the corpus digest above, not by the commit label: the digest
    # is recomputed from the checkout being published and compares actual bytes, so a map whose
    # digest matches describes this source whatever its working tree looked like when it was
    # generated.
    #
    # Neither equality with GITHUB_SHA nor a refusal of a dirty build is required here, because
    # a map is always generated before the commit that carries it. Demanding either would mean
    # generating the map inside the publishing job, or committing the code and the map
    # separately; the first is impossible (the semantic layer needs a model) and the second is
    # not this project's workflow. What is required is that the page does not hide the state it
    # was built in, which is checked above, and a caller that does generate the map in the same
    # job can still pin it exactly with -RequireCommit.
    if (-not [string]::IsNullOrWhiteSpace($RequireCommit)) {
        if ($isDirty) {
            Add-Violation ("The map was built from an uncommitted working tree ('{0}'), so it cannot " +
                           "describe commit '{1}'." -f $claimedCommit, $RequireCommit)
        }
        elseif ($claimedCommit -ne $RequireCommit) {
            Add-Violation ("The map describes commit '{0}' but '{1}' was required." -f
                           $claimedCommit, $RequireCommit)
        }
    }
}

# --- 3f. Only the expected files may be published ---------------------------
# The Pages workflow uploads the whole docs/ tree, so every file in this directory goes live. The
# content scan below only looks at five extensions, so anything else here was published without
# ever being scanned (R4-G04). The directory has one generator and a known output, so the set is
# named rather than filtered.
$expectedFiles = @('index.html', 'relationships.json', 'metadata.json')

foreach ($file in Get-ChildItem -Path $MapDirectory -Recurse -File -Force) {
    $relative = [System.IO.Path]::GetRelativePath($MapDirectory, $file.FullName).Replace('\', '/')
    if ($expectedFiles -notcontains $relative) {
        Add-Violation ("'{0}' is not part of the generated map. Only {1} are published from this " +
                       "directory; anything else goes live without being scanned." -f
                       $relative, ($expectedFiles -join ", "))
    }
}

# --- 4. Content scan --------------------------------------------------------
# The address and path patterns live in graphify/LeakScanner.psm1 so each category can be
# driven with a must-refuse and a must-allow string from the self-test, rather than being
# exercised only by publishing a map that happens to contain one (R3-G07).
Import-Module (Join-Path $PSScriptRoot "LeakScanner.psm1") -Force

$forbiddenSourceMarkers = @(
    'review-backlog',
    'project-remediation-plan',
    'AGENTS.md',
    'coverage-reports',
    'sonar-reports',
    'graphify-out',
    'CODE_QUALITY_EXCEPTIONS'
)

# Every file the Pages artifact contains, not five extensions inside one directory. The workflow
# uploads the whole docs tree (actions/upload-pages-artifact, path: ./docs), so a file this loop
# skipped went live without ever being read (R20-G01). Binary files are read as text too: an
# embedded ASCII secret matches the same patterns whatever the file's extension says it is.
$publishedRoot = Split-Path -Parent $MapDirectory
$mapPrefix = (Resolve-Path $MapDirectory).Path.TrimEnd([char]'\', [char]'/')
$strictUtf8 = [System.Text.UTF8Encoding]::new($false, $true)

foreach ($file in Get-ChildItem -Path $publishedRoot -Recurse -File -Force) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    # Which rules apply depends on what the file is (R20-G01): the generated map gets every rule;
    # a page a person wrote gets every rule but the loopback one, because it tells readers to
    # connect to localhost; bytes that are not UTF-8 text are scanned for credentials and for a
    # drive path long enough not to be chance.
    $scope = 'Binary'
    try {
        $raw = $strictUtf8.GetString($bytes)
        if ($raw.IndexOf([char]0) -lt 0) {
            $scope = if ($file.FullName.StartsWith($mapPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { 'Generated' } else { 'Authored' }
        }
    } catch [System.Text.DecoderFallbackException] {
        $raw = [System.Text.Encoding]::Latin1.GetString($bytes)
    }
    if ($scope -eq 'Binary' -and -not $raw) { $raw = [System.Text.Encoding]::Latin1.GetString($bytes) }

    foreach ($marker in $forbiddenSourceMarkers) {
        if ($raw -like "*$marker*") {
            Add-Violation "$($file.Name) references a forbidden source: $marker"
        }
    }

    foreach ($finding in (Get-LeakFinding -Text $raw -Scope $scope)) {
        Add-Violation "$($file.Name) ($scope) contains a $($finding.Kind): $($finding.Sample)"
    }
}

# --- 5. Report --------------------------------------------------------------
if ($violations.Count -gt 0) {
    Write-Host "Publish gate FAILED for '$MapDirectory':" -ForegroundColor Red
    foreach ($violation in $violations) { Write-Host "  - $violation" -ForegroundColor Red }
    exit 1
}

Write-Host "Publish gate passed for '$MapDirectory'." -ForegroundColor Green
exit 0
