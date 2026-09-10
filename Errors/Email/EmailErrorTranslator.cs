namespace AsposeMcpServer.Errors.Email;

/// <summary>
///     Translates raw Aspose.Email / BCL exceptions into sanitized BCL exceptions for the
///     Email tool surface. No inner-exception <see cref="Exception.Message" /> text is ever
///     forwarded to the emitted error — only fixed sentinels from
///     <see cref="ErrorMessageBuilder" /> flow through. Modelled after
///     <c>CellsErrorTranslator</c> so identical failure modes produce identical BCL types
///     and sanitized messages across tool families.
/// </summary>
/// <remarks>
///     The audit found no leaking Email handler catch blocks at Phase B time; this
///     translator is provided for parity so future Email handlers have a ready
///     sanitization entry-point without needing to introduce raw <c>ex.Message</c> text.
/// </remarks>
public static class EmailErrorTranslator
{
    /// <summary>
    ///     Maps an arbitrary exception thrown during an Email (Aspose.Email) operation to a
    ///     sanitized BCL exception. The mapping is:
    ///     <list type="bullet">
    ///         <item>
    ///             <see cref="UnauthorizedAccessException" /> or <see cref="DirectoryNotFoundException" /> →
    ///             <see cref="UnauthorizedAccessException" /> (output-directory context)
    ///         </item>
    ///         <item>All other → <see cref="InvalidOperationException" /></item>
    ///     </list>
    /// </summary>
    /// <param name="ex">
    ///     The raw exception thrown from an Aspose.Email API or IO call. Must not be
    ///     <c>null</c>.
    /// </param>
    /// <param name="contextBasename">
    ///     Optional sanitized basename of the file being processed. Used in
    ///     output-directory messages to add context without leaking a full path. May be
    ///     <c>null</c>.
    /// </param>
    /// <returns>
    ///     A sanitized BCL exception ready to be thrown. The returned exception is always a
    ///     new instance — the original <paramref name="ex" /> is never re-thrown or attached
    ///     as an inner exception, so internal Aspose stack frames do not escape.
    /// </returns>
    public static Exception Translate(Exception ex, string? contextBasename = null)
    {
        if (ErrorMessageBuilder.IsOutputDirectoryNotWritable(ex.Message))
            return ex;

        switch (ex)
        {
            case UnauthorizedAccessException:
            case DirectoryNotFoundException:
                return new UnauthorizedAccessException(
                    ErrorMessageBuilder.OutputDirectoryNotWritable(contextBasename));

            default:
                return new InvalidOperationException(ErrorMessageBuilder.ProcessingFailed());
        }
    }

    /// <summary>
    ///     Maps a failure raised while writing to the caller's output directory. Unlike
    ///     <see cref="Translate" />, which wraps a whole operation and therefore cannot tell an
    ///     input failure from an output one, this is called only around a write, so an
    ///     <see cref="IOException" /> here is unambiguously an output failure and is reported as
    ///     one instead of as the generic processing sentinel.
    /// </summary>
    /// <param name="ex">The exception raised by the directory creation or the file write.</param>
    /// <param name="contextBasename">
    ///     Sanitized basename of the file being written, or <c>null</c> when the failure was
    ///     creating the directory itself. Never a full path.
    /// </param>
    /// <returns>
    ///     A sanitized exception naming the output directory as the problem. The type mirrors the
    ///     original so callers that distinguish access failures from other IO still can.
    /// </returns>
    public static Exception TranslateOutputFailure(Exception ex, string? contextBasename = null)
    {
        var message = ErrorMessageBuilder.OutputDirectoryNotWritable(contextBasename);

        return ex switch
        {
            UnauthorizedAccessException => new UnauthorizedAccessException(message),
            DirectoryNotFoundException => new UnauthorizedAccessException(message),
            IOException => new IOException(message),
            _ => new InvalidOperationException(ErrorMessageBuilder.ProcessingFailed())
        };
    }
}
