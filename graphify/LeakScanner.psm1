<#
.SYNOPSIS
    Scans published text for identifiers that must never leave this machine (R3-G07).

.DESCRIPTION
    The publish gate used to hold these patterns inline, which made them impossible to
    exercise: the only way to learn whether a pattern matched was to publish a map
    containing the string. They also covered a fraction of what leaks — one private IPv4
    range was missing entirely, IPv6 was not considered, and outside `/home` and `/Users`
    an absolute path went unnoticed.

    Everything here is a pure function of its input text, so `graphify/self-test.py` can
    drive it with both a string that must be refused and a string that must be allowed for
    every category.

    Public project identifiers are removed before scanning. The published site, the
    container registry and the vendor documentation are meant to be visible, and several of
    them would otherwise match the hostname patterns.
#>

# Strings that are published on purpose and must not be read as leaked data.
$script:PublicIdentifiers = @(
    'xjustloveux\.github\.io',
    'ghcr\.io/xjustloveux',
    'github\.com/xjustloveux',
    'xjustloveux/tap',
    'sonarcloud\.io',
    'codecov\.io',
    'docs\.aspose\.com',
    'purchase\.aspose\.com',
    'releases\.aspose\.com',
    'modelcontextprotocol\.io',
    'schemas\.microsoft\.com',
    'www\.w3\.org',
    'fonts\.googleapis\.com',
    'fonts\.gstatic\.com',
    'cdnjs\.cloudflare\.com',
    'unpkg\.com'
)

# Filesystem locations. A POSIX path is only treated as absolute when it starts at one of
# the system roots: the map legitimately carries site-relative references such as
# "/architecture-map/index.html", and treating every leading slash as a path would refuse
# the page for describing itself.
$script:PathPatterns = @(
    @{ Kind = 'Windows drive path'; Pattern = '(?<![A-Za-z])[A-Za-z]:[\\/](?!/)' },
    @{ Kind = 'UNC or device path'; Pattern = '\\\\(\?|\.)\\|\\\\[A-Za-z0-9._-]+\\[A-Za-z0-9._$-]+' },
    @{ Kind = 'POSIX absolute path'
       Pattern = '/(home|Users|root|etc|var|usr|opt|srv|mnt|media|proc|sys|tmp|private)/[A-Za-z0-9_.-]' }
)

# Paths that name this machine's layout, for pages a person wrote. A how-to legitimately says
# `/usr/share/fonts`, `/tmp/output` or `C:\Tools\...`; what it must not say is where the author
# keeps their home or their checkouts (R20-G01). A placeholder such as `C:\Users\...` or
# `/home/<user>` does not begin with a letter after the root, and is allowed.
$script:AuthoredPathPatterns = @(
    @{ Kind = 'home directory path'
       Pattern = '/(home|Users|root)/[A-Za-z][A-Za-z0-9_.-]*|(?<![A-Za-z])[A-Za-z]:\\Users\\[A-Za-z][A-Za-z0-9_.-]*' },
    @{ Kind = 'workspace root path'
       Pattern = '(?i)(?<![A-Za-z])[A-Za-z]:\\(git|src|repos?|source|work|projects?|dev)\\' }
)

# The drive-path rule for bytes that were never text: it has to continue as a path to count.
$script:BinaryPathPattern = @{ Kind = 'Windows drive path'
                                Pattern = '(?<![A-Za-z])[A-Za-z]:\\[A-Za-z0-9_ .\\-]{4,}' }

# Addresses and names that only mean something inside a private network. RFC 1918 in full,
# loopback, link-local (which covers the cloud metadata address), carrier-grade NAT, and
# the IPv6 equivalents.
$script:AddressPatterns = @(
    @{ Kind = 'loopback host'; Pattern = '(?i)\blocalhost\b|\b127(\.\d{1,3}){3}\b|\b0\.0\.0\.0\b' },
    @{ Kind = 'RFC 1918 address'
       Pattern = '\b10(\.\d{1,3}){3}\b|\b192\.168(\.\d{1,3}){2}\b|\b172\.(1[6-9]|2\d|3[01])(\.\d{1,3}){2}\b' },
    @{ Kind = 'link-local or metadata address'; Pattern = '\b169\.254(\.\d{1,3}){2}\b' },
    @{ Kind = 'carrier-grade NAT address'
       Pattern = '\b100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])(\.\d{1,3}){2}\b' },
    @{ Kind = 'IPv6 loopback'; Pattern = '(?<![0-9A-Fa-f:])::1(?![0-9A-Fa-f:.])' },
    @{ Kind = 'IPv6 unique local address'; Pattern = '(?i)\b[fF][cCdD][0-9a-fA-F]{2}:[0-9a-fA-F:]{2,}' },
    @{ Kind = 'IPv6 link-local address'; Pattern = '(?i)\bfe80::[0-9a-fA-F:]*' },
    @{ Kind = 'internal hostname'; Pattern = '(?i)\b[a-z0-9][a-z0-9-]*\.internal\b' }
)

# Credentials. The first pattern here wants the keyword and the delimiter adjacent, which is
# the shape `password=` and `password:` take -- and which a quoted key does not, because the
# closing quote sits between them. The entries after it cover the shapes that gap left through
# (R19-G01): quoted JSON/YAML values, an Authorization header or a bare bearer credential, the
# well-known mount points a container reads its secrets from, and the vendor key formats whose
# prefixes identify them without any keyword nearby.
#
# Placeholders are refused, not excused (R21-G02). `"client_secret": "<your value>"` is the shape
# a credential takes, and a gate that guesses which values are examples is a gate that can be
# talked past by formatting a real one as an example. A page that must show the shape does so
# without a credential-named key.
#
# Everything here is deliberately anchored on something more than a keyword. This project's own
# published documentation is full of `ASPOSE_AUTH_APIKEY_*` variable names, `--auth-apikey-*`
# flags and an `X-API-Key` header name, none of which carry a credential; a pattern that fired on
# the word alone would refuse the documentation for describing the option.
$script:OtherPatterns = @(
    @{ Kind = 'secret-shaped string'
       Pattern = '(?i)(api[_-]?key|secret|password|passwd|bearer\s|private[_-]key|BEGIN [A-Z ]*PRIVATE KEY)\s*[:=]\s*\S|sk-[A-Za-z0-9]{16,}|gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{60,}|eyJ[A-Za-z0-9_-]{20,}\.' },
    @{ Kind = 'quoted secret assignment'
       Pattern = '(?im)["'']\b(api[_-]?key|client[_-]?secret|secret[_-]?(key|value)?|aws[_-]?secret[_-]?access[_-]?key|secret[_-]?access[_-]?key|password|passwd|pwd|access[_-]?token|refresh[_-]?token|private[_-]?key)\b["'']\s*[:=]\s*["''][^"'']{4,}["'']|(?i)^\s*(api[_-]?key|client[_-]?secret|secret[_-]?key|aws[_-]?secret[_-]?access[_-]?key|password|passwd|access[_-]?token)\s*[:=]\s*["'']?[^"''\s]{4,}["'']?' },
    @{ Kind = 'authorization credential'
       Pattern = '(?i)\bauthorization\b\s*[:=]\s*["'']?(bearer|basic|token|digest)\s+\S|(?i)\bbearer\s+[A-Za-z0-9._~+/-]{16,}={0,2}' },
    @{ Kind = 'mounted secret path'
       Pattern = '(?i)/(run|var/run|etc)/secrets?/\S|(?i)\$\{\{\s*secrets\.[A-Za-z_][A-Za-z0-9_]*\s*\}\}' },
    @{ Kind = 'vendor credential format'
       Pattern = '\b(AKIA|ASIA)[0-9A-Z]{16}\b|\bAIza[0-9A-Za-z_-]{35}\b|\bxox[abprs]-[A-Za-z0-9-]{10,}|\bnpm_[A-Za-z0-9]{36}\b|\bglpat-[A-Za-z0-9_-]{20,}' },
    # A PEM private-key envelope, header or footer. This used to be a branch of the assignment
    # pattern above, which requires `:` or `=` after the keyword; a PEM header ends in `-----`,
    # so every private key passed (R21-G01). Case-sensitive on purpose: the envelope is upper-case
    # by specification, and a prose mention of "private key" is not one.
    @{ Kind = 'private key envelope'
       Pattern = '-----(BEGIN|END) (?:[A-Z0-9]+ )*PRIVATE KEY-----' },
    @{ Kind = 'e-mail address'; Pattern = '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}' }
)


function Remove-PublicIdentifier {
    <#
    .SYNOPSIS
        Blanks the identifiers this project publishes on purpose.

    .PARAMETER Text
        Text to scrub.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    foreach ($pattern in $script:PublicIdentifiers) {
        $Text = [regex]::Replace($Text, $pattern, '')
    }
    return $Text
}

function Get-LeakFinding {
    <#
    .SYNOPSIS
        Returns every private identifier found in the given text.

    .PARAMETER Text
        Text to scan; public project identifiers are removed first.

    .OUTPUTS
        One object per category that matched, carrying the category name and the matched
        text truncated to forty characters.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        # What kind of file the text came from (R20-G01). The rules are about what must not leave
        # this machine; which of them apply depends on whether the file is a generated artifact,
        # a page a person wrote for readers, or bytes that were never text at all.
        [ValidateSet('Generated', 'Authored', 'Binary')][string]$Scope = 'Generated'
    )

    $scrubbed = Remove-PublicIdentifier -Text $Text
    $findings = @()

    $checks = switch ($Scope) {
        'Generated' { $script:PathPatterns + $script:AddressPatterns + $script:OtherPatterns }
        # A page telling the reader to connect to localhost or bind 0.0.0.0 documents the product.
        # Every other rule still applies: a private address or a drive path in prose is a leak.
        'Authored'  { $script:AuthoredPathPatterns + ($script:AddressPatterns | Where-Object { $_.Kind -ne 'loopback host' }) + $script:OtherPatterns }
        # Bytes read as text. Short shapes match by chance; a drive path counts only when it goes
        # on as one, and addresses are not looked for at all. Credentials keep their full weight:
        # an ASCII key in image metadata is the case this scope exists to catch.
        'Binary'    { @($script:BinaryPathPattern) + $script:OtherPatterns }
    }

    foreach ($check in $checks) {
        $match = [regex]::Match($scrubbed, $check.Pattern)
        if (-not $match.Success) { continue }

        $sample = $match.Value
        if ($sample.Length -gt 40) { $sample = $sample.Substring(0, 40) }
        $findings += [PSCustomObject]@{ Kind = $check.Kind; Sample = $sample }
    }

    return $findings
}

Export-ModuleMember -Function Remove-PublicIdentifier, Get-LeakFinding
