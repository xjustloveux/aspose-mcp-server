using System.Text;
using System.Text.RegularExpressions;

namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Single source of truth for sanitizing attacker-controlled filenames carried on
///     OLE objects (Word <c>OlePackage.FileName</c>, Cells <c>OleObject.SourceFullName</c>,
///     Slides <c>OleObjectFrame.EmbeddedData.EmbeddedFileName</c>) before writing to disk,
///     and for neutralizing log/error-message interpolation sites that touch those values.
///     Consolidates all OLE-specific hardening (BiDi / control / UTF-8 byte clamp / Windows
///     reserved names / path-traversal neutralization) so the rules live in exactly one place
///     and stay consistent across the Word, Excel, and PowerPoint tools (AC-18).
/// </summary>
public static class OleSanitizerHelper
{
    /// <summary>
    ///     Maximum UTF-8 byte length accepted on disk (ext4 <c>NAME_MAX</c> / NTFS component limit).
    /// </summary>
    private const int MaxFileNameBytes = 255;

    /// <summary>
    ///     Match timeout applied to every <see cref="Regex" /> in this class. 100 ms is
    ///     ample for any legitimate filename (ext4 NAME_MAX = 255 bytes); if a crafted
    ///     input ever manages to stall the engine beyond this bound the call site receives
    ///     a <see cref="System.Text.RegularExpressions.RegexMatchTimeoutException" />,
    ///     which propagates as an unhandled exception and aborts the tool invocation
    ///     rather than tying up the thread indefinitely (DoS / ReDoS defense-in-depth).
    /// </summary>
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(100);

    /// <summary>
    ///     Regex stripping BiDi override code points (U+200E, U+200F, U+202A-U+202E, U+2066-U+2069).
    ///     RTLO attacks (e.g. <c>photo\u202egpj.exe</c>) are defeated here prior to other passes.
    /// </summary>
    private static readonly Regex BiDiOverrides =
        new("[\u200E\u200F\u202A-\u202E\u2066-\u2069]", RegexOptions.Compiled, RegexTimeout);

    /// <summary>
    ///     Regex stripping C0 controls (U+0000-U+001F) and C1 controls (U+007F-U+009F).
    /// </summary>
    private static readonly Regex ControlChars =
        new("[\u0000-\u001F\u007F-\u009F]", RegexOptions.Compiled, RegexTimeout);

    /// <summary>
    ///     Regex matching ANSI CSI escape sequences. Strips terminal-control injection
    ///     in log/error content (F-4).
    /// </summary>
    private static readonly Regex AnsiCsi =
        new("\u001B\\[[0-9;]*[A-Za-z]", RegexOptions.Compiled, RegexTimeout);

    /// <summary>
    ///     Regex matching trailing whitespace / dot characters that are stripped at the end
    ///     of a candidate filename (space, tab, form-feed, vertical tab, NBSP, dot).
    /// </summary>
    private static readonly Regex TrailingJunk =
        new("[\\s.\u00A0\u000B\u000C]+$", RegexOptions.Compiled, RegexTimeout);

    /// <summary>
    ///     Windows reserved device names. Kept cross-platform so a file authored on Linux
    ///     is still safe if later shipped to a Windows MCP client.
    /// </summary>
    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
    };

    /// <summary>
    ///     Sanitizes an OLE-derived raw filename into a disk-safe form.
    /// </summary>
    /// <param name="rawName">
    ///     Raw filename carried on the OLE payload. May be <c>null</c>, empty, entirely
    ///     composed of stripped characters, or attacker-controlled (traversal, UNC,
    ///     absolute-path smuggling, RTLO, NUL, CRLF, ANSI CSI, etc.).
    /// </param>
    /// <param name="index">
    ///     Zero-based OLE index within its container — used only for the fallback
    ///     <c>ole_&lt;index&gt;&lt;extension&gt;</c> name when the raw name is empty
    ///     after sanitization.
    /// </param>
    /// <param name="progId">
    ///     ProgId reported by the container (e.g. <c>Excel.Sheet.12</c>). Used only to pick
    ///     a fallback extension when the raw name is unusable. May be <c>null</c>.
    /// </param>
    /// <returns>
    ///     Tuple of <c>(suggested, sanitizedFromRaw)</c> where <c>suggested</c> is a disk-safe
    ///     filename (never null, never empty, never containing path separators, at most 255
    ///     UTF-8 bytes), and <c>sanitizedFromRaw</c> is <c>true</c> when the returned value
    ///     differs from <paramref name="rawName" />.
    /// </returns>
    /// <remarks>
    ///     Order of operations (F-1): strip BiDi → strip C0/C1 controls → strip path-component
    ///     prefixes → replace separators/colon/NUL with <c>_</c> → collapse <c>..</c> → trim
    ///     trailing whitespace/dots → reserved-name prefix → empty-fallback → UTF-8 byte clamp.
    ///     Idempotent: <c>Sanitize(Sanitize(x)) == Sanitize(x)</c>. Unicode normalization is
    ///     deliberately NOT applied to keep idempotence guaranteed.
    /// </remarks>
    /// <exception cref="System.Text.RegularExpressions.RegexMatchTimeoutException">
    ///     Thrown if any regex pass exceeds <see cref="RegexTimeout" /> (100 ms). Indicates a
    ///     pathological input; callers should treat this as an untrusted-input rejection signal.
    /// </exception>
    public static (string suggested, bool sanitizedFromRaw) SanitizeOleFileName(
        string? rawName, int index, string? progId)
    {
        var original = rawName ?? string.Empty;
        var working = original;

        working = BiDiOverrides.Replace(working, string.Empty);
        working = ControlChars.Replace(working, string.Empty);

        if (working.Length >= 2 && char.IsLetter(working[0]) && working[1] == ':')
            working = working[2..];

        while (working.StartsWith("\\\\", StringComparison.Ordinal) ||
               working.StartsWith("//", StringComparison.Ordinal))
            working = working[2..];

        working = working.TrimStart('/', '\\');

        var lastSeparator = working.LastIndexOfAny(['/', '\\']);
        if (lastSeparator >= 0)
            working = working[(lastSeparator + 1)..];

        working = working.Replace(":", "_", StringComparison.Ordinal);
        working = working.Replace("\0", "_", StringComparison.Ordinal);

        foreach (var invalid in Path.GetInvalidFileNameChars())
        {
            if (invalid == '/' || invalid == '\\' || invalid == ':' || invalid == '\0') continue;
            working = working.Replace(invalid.ToString(), "_", StringComparison.Ordinal);
        }

        while (working.Contains("..", StringComparison.Ordinal))
            working = working.Replace("..", string.Empty, StringComparison.Ordinal);

        working = TrailingJunk.Replace(working, string.Empty).TrimStart();

        if (working.Length > 0)
        {
            var stem = Path.GetFileNameWithoutExtension(working);
            if (ReservedNames.Contains(stem))
                working = "_" + working;
        }

        if (string.IsNullOrEmpty(working))
            working = $"ole_{index}{ExtensionFromProgId(progId)}";

        working = ClampToUtf8ByteLimit(working);

        return (working, !string.Equals(working, rawName, StringComparison.Ordinal));
    }

    /// <summary>
    ///     Maps a ProgId to a conventional extension. Covers the common Office / Adobe /
    ///     generic-package ProgIds; falls back to <c>.bin</c> for unknown strings.
    /// </summary>
    /// <param name="progId">
    ///     ProgId string as reported by the source container (e.g. <c>Excel.Sheet.12</c>,
    ///     <c>Word.Document.12</c>, <c>PowerPoint.Show.12</c>, <c>AcroExch.Document</c>,
    ///     <c>Package</c>). Comparison is case-insensitive and tolerates the legacy
    ///     lowercase variants Aspose.Slides emits (e.g. <c>excel.sheet.12</c>).
    /// </param>
    /// <returns>
    ///     Extension with a leading dot (e.g. <c>".xlsx"</c>) or <c>".bin"</c> for an
    ///     unrecognized / null / empty ProgId.
    /// </returns>
    public static string ExtensionFromProgId(string? progId)
    {
        if (string.IsNullOrWhiteSpace(progId)) return ".bin";

        var p = progId.Trim().ToLowerInvariant();

        if (p.StartsWith("excel.sheet.12", StringComparison.Ordinal)) return ".xlsx";
        if (p.StartsWith("excel.sheetmacroenabled.12", StringComparison.Ordinal)) return ".xlsm";
        if (p.StartsWith("excel.sheetbinary.12", StringComparison.Ordinal)) return ".xlsb";
        if (p.StartsWith("excel.sheet.8", StringComparison.Ordinal)) return ".xls";
        if (p.StartsWith("excel.sheet", StringComparison.Ordinal)) return ".xls";
        if (p.StartsWith("excel.chart", StringComparison.Ordinal)) return ".xls";

        if (p.StartsWith("word.document.12", StringComparison.Ordinal)) return ".docx";
        if (p.StartsWith("word.documentmacroenabled.12", StringComparison.Ordinal)) return ".docm";
        if (p.StartsWith("word.document", StringComparison.Ordinal)) return ".doc";

        if (p.StartsWith("powerpoint.show.12", StringComparison.Ordinal)) return ".pptx";
        if (p.StartsWith("powerpoint.showmacroenabled.12", StringComparison.Ordinal)) return ".pptm";
        if (p.StartsWith("powerpoint.show", StringComparison.Ordinal)) return ".ppt";
        if (p.StartsWith("powerpoint.slide", StringComparison.Ordinal)) return ".ppt";

        if (p.StartsWith("acroexch.document", StringComparison.Ordinal)) return ".pdf";

        return ".bin";
    }

    /// <summary>
    ///     Normalizes an extension so the returned value always has a leading dot and
    ///     contains no path separators. Handles the Aspose.Slides round-trip quirk where
    ///     <c>EmbeddedFileExtension</c> may come back with or without a leading dot.
    /// </summary>
    /// <param name="extension">
    ///     Raw extension from an Aspose API (e.g. <c>"xlsx"</c>, <c>".xlsx"</c>,
    ///     <c>null</c>, or <c>string.Empty</c>).
    /// </param>
    /// <returns>
    ///     Normalized extension such as <c>".xlsx"</c>, or <c>string.Empty</c> when the
    ///     input is null / empty after trimming.
    /// </returns>
    public static string NormalizeExtension(string? extension)
    {
        if (string.IsNullOrWhiteSpace(extension)) return string.Empty;
        var trimmed = extension.Trim();
        trimmed = trimmed.Replace('\\', '_').Replace('/', '_');
        if (trimmed.Length == 0) return string.Empty;
        return trimmed.StartsWith('.') ? trimmed : "." + trimmed;
    }

    /// <summary>
    ///     Determines whether the supplied name matches a Windows reserved device name,
    ///     case-insensitive, ignoring any extension suffix.
    /// </summary>
    /// <param name="candidate">
    ///     Candidate filename (with or without extension). <c>null</c> / empty returns <c>false</c>.
    /// </param>
    /// <returns>
    ///     <c>true</c> when the stem (part before the final <c>.</c>) is a Windows reserved
    ///     device name (CON / PRN / AUX / NUL / COM1-9 / LPT1-9); otherwise <c>false</c>.
    /// </returns>
    public static bool IsWindowsReservedName(string? candidate)
    {
        if (string.IsNullOrEmpty(candidate)) return false;
        var stem = Path.GetFileNameWithoutExtension(candidate);
        return !string.IsNullOrEmpty(stem) && ReservedNames.Contains(stem);
    }

    /// <summary>
    ///     Strips log-injection vectors from attacker-controlled text prior to interpolation
    ///     into any log template, error message, or diagnostic string (F-4). Removes CR, LF,
    ///     TAB, NUL, C0 (U+0000-U+001F) and C1 (U+007F-U+009F) control characters, and ANSI
    ///     CSI escape sequences. Callers MUST route <c>rawFileName</c>, <c>path</c>,
    ///     <c>progId</c>, <c>LinkTarget</c>, and any inner-exception fragment through this
    ///     helper before <see cref="Microsoft.Extensions.Logging.ILogger" /> sites or
    ///     error-message construction.
    /// </summary>
    /// <param name="value">
    ///     Raw string from an attacker-reachable surface. <c>null</c> returns
    ///     <see cref="string.Empty" />.
    /// </param>
    /// <returns>A log-safe rendering (never <c>null</c>; empty input → empty string).</returns>
    /// <exception cref="System.Text.RegularExpressions.RegexMatchTimeoutException">
    ///     Thrown if any regex pass exceeds <see cref="RegexTimeout" /> (100 ms). Indicates a
    ///     pathological input; callers should treat this as an untrusted-input rejection signal.
    /// </exception>
    public static string SanitizeForLog(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var stripped = ControlChars.Replace(value, string.Empty);
        stripped = AnsiCsi.Replace(stripped, string.Empty);
        return stripped;
    }

    /// <summary>
    ///     Clamps a string to at most <see cref="MaxFileNameBytes" /> UTF-8 bytes by trimming
    ///     from the right. Correct for ext4 (255-byte <c>NAME_MAX</c>) and conservative but
    ///     safe for NTFS / HFS+ / APFS (per design §8).
    /// </summary>
    /// <param name="name">Non-null candidate name (already validated non-empty).</param>
    /// <returns>A string whose UTF-8 byte count is ≤ 255.</returns>
    private static string ClampToUtf8ByteLimit(string name)
    {
        if (Encoding.UTF8.GetByteCount(name) <= MaxFileNameBytes) return name;

        var extension = Path.GetExtension(name);
        var extBytes = string.IsNullOrEmpty(extension) ? 0 : Encoding.UTF8.GetByteCount(extension);
        var stem = string.IsNullOrEmpty(extension) ? name : name[..^extension.Length];

        if (extBytes >= MaxFileNameBytes)
        {
            var truncated = extension;
            while (truncated.Length > 0 && Encoding.UTF8.GetByteCount(truncated) > MaxFileNameBytes)
                truncated = truncated[..^1];
            return truncated;
        }

        while (stem.Length > 0 && Encoding.UTF8.GetByteCount(stem) + extBytes > MaxFileNameBytes)
            stem = stem[..^1];

        return stem + extension;
    }
}
