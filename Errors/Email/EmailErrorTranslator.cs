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
}
