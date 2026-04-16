using AsposeMcpServer.Helpers.Ole;

namespace AsposeMcpServer.Errors.Ole;

/// <summary>
///     Central builder for sanitized OLE-error messages (F-10, F-11). Every
///     attacker-reachable field is routed through
///     <see cref="OleSanitizerHelper.SanitizeForLog" /> and then inserted into a fixed
///     template. Raw string concatenation of user input is forbidden (code-reviewer gate).
///     Messages are byte-identical across the Word / Excel / PowerPoint tools for
///     equivalent failure modes (AC-19).
/// </summary>
public static class OleErrorMessageBuilder
{
    /// <summary>
    ///     Fixed sentinel emitted for invalid-password cases (F-5). Value is byte-identical
    ///     to <see cref="ErrorMessageBuilder.InvalidPassword" /> so that OLE callers that
    ///     reference this constant directly do not observe a behavioral change.
    /// </summary>
    public const string InvalidPasswordSentinel =
        "The source file requires a password, or the supplied password is incorrect.";

    /// <summary>Fixed sentinel emitted for linked-OLE extraction attempts.</summary>
    public const string LinkedCannotExtractSentinel =
        "The requested OLE object is a link and has no embedded payload; extraction is skipped.";

    /// <summary>
    ///     Fixed sentinel emitted when the translator cannot recover the real index / count
    ///     from the underlying exception (replaces the previous <c>-1 / 0</c> leak).
    /// </summary>
    public const string IndexOutOfRangeUnknownSentinel =
        "The supplied OLE index is out of range.";

    /// <summary>
    ///     Builds the <c>unknown-operation</c> message — includes the sanitized operation
    ///     token and the supported enum list; never the inner exception or any path.
    /// </summary>
    /// <param name="operation">The rejected operation token (may be null / attacker-controlled).</param>
    /// <returns>Sanitized error message.</returns>
    public static string UnknownOperation(string? operation)
    {
        var safe = OleSanitizerHelper.SanitizeForLog(operation);
        return $"Unknown OLE operation '{safe}'. Supported: list, extract, extract_all, remove.";
    }

    /// <summary>
    ///     Builds the <c>invalid-path</c> message — fixed sentinel plus the sanitized
    ///     basename only; never the full path or allowlist.
    /// </summary>
    /// <param name="pathFragment">Attacker-controlled path (may be null).</param>
    /// <returns>Sanitized error message.</returns>
    public static string InvalidPath(string? pathFragment)
    {
        var basename = string.IsNullOrEmpty(pathFragment) ? string.Empty : Path.GetFileName(pathFragment);
        var safe = OleSanitizerHelper.SanitizeForLog(basename);
        return string.IsNullOrEmpty(safe)
            ? "The supplied path is invalid or outside the configured allowlist."
            : $"The supplied path '{safe}' is invalid or outside the configured allowlist.";
    }

    /// <summary>
    ///     Builds the <c>ole-index-out-of-range</c> message. When <paramref name="index" />
    ///     and <paramref name="count" /> are both known the concrete numerics are emitted;
    ///     when either is <c>null</c> (translator path where the raw exception did not
    ///     carry the values) the fixed sentinel is emitted instead so misleading
    ///     <c>-1 / 0</c> numerics never reach the caller.
    /// </summary>
    /// <param name="index">Requested index (as supplied by the caller), or <c>null</c> when unknown.</param>
    /// <param name="count">Actual OLE object count in the container snapshot, or <c>null</c> when unknown.</param>
    /// <returns>Sanitized error message.</returns>
    public static string IndexOutOfRange(int? index, int? count)
    {
        return index is null || count is null
            ? IndexOutOfRangeUnknownSentinel
            : $"OLE index {index.Value} is out of range. Container holds {count.Value} OLE object(s).";
    }

    /// <summary>
    ///     Builds the <c>output-directory-not-writable</c> message — fixed sentinel plus
    ///     sanitized basename; never the full path or parent chain.
    /// </summary>
    /// <param name="pathFragment">The rejected output directory path.</param>
    /// <returns>Sanitized error message.</returns>
    public static string OutputDirectoryNotWritable(string? pathFragment)
    {
        var basename = string.IsNullOrEmpty(pathFragment) ? string.Empty : Path.GetFileName(pathFragment);
        var safe = OleSanitizerHelper.SanitizeForLog(basename);
        return string.IsNullOrEmpty(safe)
            ? "The output directory cannot be created or is not writable."
            : $"The output directory '{safe}' cannot be created or is not writable.";
    }

    /// <summary>
    ///     Builds the <c>ole-save-failed</c> message — sanitized basename plus fixed
    ///     failure code; never the full path or inner-exception text.
    /// </summary>
    /// <param name="fileName">The sanitized (or attacker-controlled) filename in question.</param>
    /// <returns>Sanitized error message.</returns>
    public static string SaveFailed(string? fileName)
    {
        var basename = string.IsNullOrEmpty(fileName) ? string.Empty : Path.GetFileName(fileName);
        var safe = OleSanitizerHelper.SanitizeForLog(basename);
        return string.IsNullOrEmpty(safe)
            ? "Failed to write the extracted OLE payload to disk."
            : $"Failed to write the extracted OLE payload '{safe}' to disk.";
    }

    /// <summary>
    ///     Builds the <c>ole-remove-failed</c> message — fixed sentinel plus sanitized
    ///     basename; never the Aspose inner-exception text.
    /// </summary>
    /// <param name="fileName">The sanitized source basename in question.</param>
    /// <returns>Sanitized error message.</returns>
    public static string RemoveFailed(string? fileName)
    {
        var basename = string.IsNullOrEmpty(fileName) ? string.Empty : Path.GetFileName(fileName);
        var safe = OleSanitizerHelper.SanitizeForLog(basename);
        return string.IsNullOrEmpty(safe)
            ? "Failed to remove the OLE object or re-save the container."
            : $"Failed to remove the OLE object or re-save the container '{safe}'.";
    }

    /// <summary>Fixed sentinel message for the linked-OLE case.</summary>
    /// <returns>The fixed sentinel string.</returns>
    public static string LinkedCannotExtract()
    {
        return LinkedCannotExtractSentinel;
    }

    /// <summary>
    ///     Fixed sentinel message for the invalid-password case (F-5). Delegates to
    ///     <see cref="ErrorMessageBuilder.InvalidPassword" /> to keep the shared sentinel in
    ///     one place; callers observe no behavioral change.
    /// </summary>
    /// <returns>The fixed sentinel string, byte-identical to <see cref="InvalidPasswordSentinel" />.</returns>
    public static string InvalidPassword()
    {
        return ErrorMessageBuilder.InvalidPassword();
    }

    /// <summary>
    ///     Builds the <c>unsupported-legacy-format</c> message — sanitized extension token
    ///     only.
    /// </summary>
    /// <param name="extension">Extension reported by the caller (e.g. <c>".doc"</c>).</param>
    /// <returns>Sanitized error message.</returns>
    public static string UnsupportedLegacyFormat(string? extension)
    {
        var safe = OleSanitizerHelper.SanitizeForLog(extension);
        return string.IsNullOrEmpty(safe)
            ? "The legacy container format could not be opened as an OLE-capable document."
            : $"The legacy container format '{safe}' could not be opened as an OLE-capable document.";
    }
}
