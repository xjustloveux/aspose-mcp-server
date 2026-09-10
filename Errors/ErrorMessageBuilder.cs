using System.Diagnostics.CodeAnalysis;

namespace AsposeMcpServer.Errors;

/// <summary>
///     Unified hub for fixed error-message sentinels shared across all per-family translators
///     (Cells, Word, PDF, PowerPoint, Email). Every method returns a hard-coded string so that
///     raw Aspose / BCL exception text never reaches the MCP caller. Per-family translators
///     delegate to this class for cross-family parity; they may add family-specific methods
///     that live in their own builder files.
/// </summary>
public static class ErrorMessageBuilder
{
    /// <summary>Opening words shared by both forms of the output-directory message.</summary>
    private const string OutputDirectoryPrefix = "The output directory";

    /// <summary>Closing words shared by both forms of the output-directory message.</summary>
    private const string OutputDirectorySuffix = "cannot be created or is not writable.";

    /// <summary>
    ///     Returns the fixed sentinel for password failures: either the file requires a
    ///     password that was not supplied, or the supplied password is wrong.
    /// </summary>
    /// <returns>
    ///     A fixed string that never contains file-path, stack-trace, or inner-exception
    ///     text.
    /// </returns>
    [SuppressMessage("SonarAnalyzer.CSharp", "S3400",
        Justification = "Intentional method form. ErrorMessageBuilder is a builder API whose " +
                        "sibling members accept parameters (OperationFailed(verb), UnsupportedFormat(extension), " +
                        "IndexOutOfRange(itemName, rangeHint)). Exposing some sentinels as methods and " +
                        "others as fields would break call-site consistency across 27+ callers.")]
    public static string InvalidPassword()
    {
        return "The source file requires a password, or the supplied password is incorrect.";
    }

    /// <summary>
    ///     Returns a sentinel for index-out-of-range failures. When <paramref name="itemName" />
    ///     and <paramref name="rangeHint" /> are both supplied the message includes them;
    ///     otherwise a generic sentinel is returned so that unknown internal counts never
    ///     leak.
    /// </summary>
    /// <param name="itemName">
    ///     Optional human-readable name for the indexed item (e.g. "worksheet", "row").
    ///     Pass <c>null</c> or empty to get the generic form.
    /// </param>
    /// <param name="rangeHint">
    ///     Optional description of the valid range (e.g. "0–9"). Pass <c>null</c> or empty
    ///     to omit the range clause from the message.
    /// </param>
    /// <returns>
    ///     A sanitized error string. Either
    ///     <c>"The supplied index is out of range."</c> (generic) or
    ///     <c>"The supplied {itemName} index is out of range. Valid range: {rangeHint}."</c>
    ///     (specific).
    /// </returns>
    public static string IndexOutOfRange(string? itemName = null, string? rangeHint = null)
    {
        var hasItem = !string.IsNullOrWhiteSpace(itemName);
        var hasRange = !string.IsNullOrWhiteSpace(rangeHint);

        if (!hasItem && !hasRange)
            return "The supplied index is out of range.";

        var itemPart = hasItem ? $"{itemName} " : string.Empty;
        var rangePart = hasRange ? $" Valid range: {rangeHint}." : ".";
        return $"The supplied {itemPart}index is out of range.{rangePart}";
    }

    /// <summary>
    ///     Returns a sentinel for output-directory permission / existence failures. When
    ///     <paramref name="contextBasename" /> is supplied it is included as a context hint
    ///     (basename only — never a full path).
    /// </summary>
    /// <param name="contextBasename">
    ///     Optional sanitized basename of the output file or directory. Pass <c>null</c> or
    ///     empty to get the generic form. Must already be a basename; this method does NOT
    ///     strip path components.
    /// </param>
    /// <returns>
    ///     A sanitized error string. Either
    ///     <c>"The output directory cannot be created or is not writable."</c> (generic) or
    ///     <c>"The output directory for '{contextBasename}' cannot be created or is not writable."</c>
    ///     (specific).
    /// </returns>
    public static string OutputDirectoryNotWritable(string? contextBasename = null)
    {
        return string.IsNullOrWhiteSpace(contextBasename)
            ? OutputDirectoryPrefix + " " + OutputDirectorySuffix
            : $"{OutputDirectoryPrefix} for '{contextBasename}' {OutputDirectorySuffix}";
    }

    /// <summary>
    ///     Whether a message is one this class produced for an output-directory failure.
    ///     <para>
    ///         The tool-call error filter replaces exception text it does not recognise, because raw
    ///         BCL text carries absolute paths. That also discarded the authored output-directory
    ///         message, so a caller whose output directory was unusable was told the input file could
    ///         not be found, or given the generic processing sentinel. This lets the filter tell the
    ///         two apart. Only text built here matches: the fixed opening and closing words must both
    ///         be present, and the only variable part is a basename the caller already supplied.
    ///     </para>
    /// </summary>
    /// <param name="message">The exception message to classify.</param>
    /// <returns><c>true</c> when the message was produced by <see cref="OutputDirectoryNotWritable" />.</returns>
    public static bool IsOutputDirectoryNotWritable(string? message)
    {
        if (message == null) return false;
        if (string.Equals(message, OutputDirectoryNotWritable(), StringComparison.Ordinal)) return true;

        // Matching only the opening and closing words let anything sit between them, and whatever
        // sat there was forwarded to the caller verbatim — an exception raised deep inside a
        // library, carrying a full path, needed only to begin and end the right way to be treated
        // as authored text (R3-C09). The message is now rebuilt from a basename that has to look
        // like a basename, and only an exact match is accepted.
        const string opening = OutputDirectoryPrefix + " for '";
        if (!message.StartsWith(opening, StringComparison.Ordinal)) return false;

        var closing = message.IndexOf("' " + OutputDirectorySuffix, opening.Length, StringComparison.Ordinal);
        if (closing < 0) return false;

        var candidate = message[opening.Length..closing];
        return IsSafeBasename(candidate)
               && string.Equals(message, OutputDirectoryNotWritable(candidate), StringComparison.Ordinal);
    }

    /// <summary>
    ///     Whether a string is a bare file name, and so safe to repeat back to the caller.
    /// </summary>
    /// <param name="value">The candidate basename.</param>
    /// <returns><c>true</c> when it names a file without disclosing where the file is.</returns>
    private static bool IsSafeBasename(string value)
    {
        return value.Length is > 0 and <= 128
               && value.IndexOfAny(['/', '\\', ':', '\'', '\r', '\n']) < 0
               && value != "."
               && value != "..";
    }

    /// <summary>
    ///     Returns a sentinel for a cell range the caller supplied that could not be read.
    /// </summary>
    /// <param name="range">
    ///     The range as the caller wrote it. Only the caller's own input is repeated back, and it
    ///     is truncated so a long value cannot be used to pad the response.
    /// </param>
    /// <returns>A sanitized string naming the range that could not be used.</returns>
    public static string InvalidRange(string? range)
    {
        if (string.IsNullOrWhiteSpace(range))
            return "The supplied cell range could not be read.";

        var shown = range.Length > 64 ? range[..64] + "..." : range;
        return $"The cell range '{shown}' could not be read. Use a range such as A1:C10.";
    }

    /// <summary>
    ///     Returns a fixed sentinel for generic document-processing failures, using the
    ///     supplied <paramref name="verb" /> to make the message contextual without embedding
    ///     internal exception text.
    /// </summary>
    /// <param name="verb">
    ///     A verb phrase describing the attempted operation (e.g. <c>"process"</c>,
    ///     <c>"render"</c>, <c>"export"</c>). Defaults to <c>"process"</c>. Must be a safe
    ///     hard-coded constant — do not pass attacker-controlled input here.
    /// </param>
    /// <returns>
    ///     A sanitized string of the form <c>"Failed to {verb} the document."</c>.
    /// </returns>
    public static string OperationFailed(string verb = "process")
    {
        return $"Failed to {verb} the document.";
    }

    /// <summary>
    ///     Returns the fixed sentinel for internal processing failures where the root cause
    ///     is unknown or should not be disclosed to the caller. Callers should prefer
    ///     <see cref="OperationFailed" /> when a meaningful verb is available.
    /// </summary>
    /// <returns>
    ///     The fixed string
    ///     <c>"An internal processing error occurred. Check inputs and retry."</c>.
    /// </returns>
    [SuppressMessage("SonarAnalyzer.CSharp", "S3400",
        Justification = "Intentional method form. ErrorMessageBuilder is a builder API whose " +
                        "sibling members accept parameters (OperationFailed(verb), UnsupportedFormat(extension), " +
                        "IndexOutOfRange(itemName, rangeHint)). Exposing some sentinels as methods and " +
                        "others as fields would break call-site consistency across 18+ callers.")]
    public static string ProcessingFailed()
    {
        return "An internal processing error occurred. Check inputs and retry.";
    }

    /// <summary>
    ///     Returns a sanitized sentinel for extension unavailability. The
    ///     <paramref name="reasonTag" /> must be a controlled, hard-coded tag (e.g.
    ///     <c>"handshake-timeout"</c>, <c>"initialization-failed"</c>) — never attacker input
    ///     or raw exception text.
    /// </summary>
    /// <param name="reasonTag">
    ///     A short, controlled tag describing why the extension is unavailable. Must not
    ///     contain file paths, exception messages, or user-controlled content.
    /// </param>
    /// <returns>
    ///     A sanitized string of the form
    ///     <c>"Extension is unavailable: {reasonTag}."</c>.
    /// </returns>
    public static string ExtensionUnavailable(string reasonTag)
    {
        return $"Extension is unavailable: {reasonTag}.";
    }

    /// <summary>
    ///     Returns a sentinel for unsupported file-format failures. When
    ///     <paramref name="extension" /> is supplied it is included to aid the caller in
    ///     understanding which format was rejected.
    /// </summary>
    /// <param name="extension">
    ///     Optional file extension reported by the caller (e.g. <c>".xls"</c>). Pass
    ///     <c>null</c> or empty for the generic form. Must be a short token — this method
    ///     does NOT sanitize the value, so callers should pass only known-safe extension
    ///     strings.
    /// </param>
    /// <returns>
    ///     A sanitized string. Either
    ///     <c>"The supplied file format is not supported."</c> (generic) or
    ///     <c>"The file format '{extension}' is not supported."</c> (specific).
    /// </returns>
    public static string UnsupportedFormat(string? extension = null)
    {
        return string.IsNullOrWhiteSpace(extension)
            ? "The supplied file format is not supported."
            : $"The file format '{extension}' is not supported.";
    }

    /// <summary>
    ///     Returns the fixed sentinel for style-application failures in Word documents.
    ///     When <paramref name="styleName" /> is supplied it is included to help the caller
    ///     identify which style was rejected; the style name is treated as a safe label
    ///     (not Aspose exception text) and is never derived from <c>ex.Message</c>.
    /// </summary>
    /// <param name="styleName">
    ///     Optional name of the style that could not be applied (e.g. <c>"Heading 1"</c>).
    ///     Pass <c>null</c> or empty for the generic form. Must be a caller-supplied label,
    ///     not exception-derived text.
    /// </param>
    /// <returns>
    ///     A sanitized string. Either
    ///     <c>"Failed to apply the requested style."</c> (generic) or
    ///     <c>"Failed to apply style '{styleName}'."</c> (specific).
    /// </returns>
    public static string StyleApplicationFailed(string? styleName = null)
    {
        return string.IsNullOrWhiteSpace(styleName)
            ? "Failed to apply the requested style."
            : $"Failed to apply style '{styleName}'.";
    }

    /// <summary>
    ///     Returns the fixed sentinel for chart-operation failures in PowerPoint
    ///     presentations. When <paramref name="operation" /> is supplied it is included to
    ///     help the caller identify which chart operation failed; the value must be a safe
    ///     hard-coded constant — not attacker-controlled or exception-derived text.
    /// </summary>
    /// <param name="operation">
    ///     Optional label describing the attempted chart operation (e.g. <c>"set title"</c>,
    ///     <c>"change type"</c>). Pass <c>null</c> or empty for the generic form. Must be
    ///     a known-safe constant.
    /// </param>
    /// <returns>
    ///     A sanitized string. Either
    ///     <c>"Failed to perform the chart operation."</c> (generic) or
    ///     <c>"Failed to {operation} on the chart."</c> (specific).
    /// </returns>
    public static string ChartOperationFailed(string? operation = null)
    {
        return string.IsNullOrWhiteSpace(operation)
            ? "Failed to perform the chart operation."
            : $"Failed to {operation} on the chart.";
    }
}
