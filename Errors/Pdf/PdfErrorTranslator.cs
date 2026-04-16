using Aspose.Pdf;

namespace AsposeMcpServer.Errors.Pdf;

/// <summary>
///     Translates raw Aspose.PDF / BCL exceptions into sanitized BCL exceptions for the
///     PDF tool surface. No inner-exception <see cref="Exception.Message" /> text is ever
///     forwarded to the emitted error — only fixed sentinels from
///     <see cref="ErrorMessageBuilder" /> or
///     <see cref="PdfErrorMessageBuilder" /> flow through. Modelled after
///     <c>CellsErrorTranslator</c> so identical failure modes produce identical BCL types
///     and sanitized messages across tool families.
/// </summary>
public static class PdfErrorTranslator
{
    /// <summary>
    ///     Maps an arbitrary exception thrown during a PDF (Aspose.PDF) operation to a
    ///     sanitized BCL exception. The mapping is:
    ///     <list type="bullet">
    ///         <item><see cref="Aspose.Pdf.InvalidPasswordException" /> → <see cref="UnauthorizedAccessException" /></item>
    ///         <item>
    ///             <see cref="UnauthorizedAccessException" /> or <see cref="DirectoryNotFoundException" /> →
    ///             <see cref="UnauthorizedAccessException" /> (output-directory context)
    ///         </item>
    ///         <item>All other → <see cref="InvalidOperationException" /></item>
    ///     </list>
    /// </summary>
    /// <param name="ex">
    ///     The raw exception thrown from an Aspose.PDF API or IO call. Must not be
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
            case InvalidPasswordException:
                return new UnauthorizedAccessException(ErrorMessageBuilder.InvalidPassword());

            case UnauthorizedAccessException:
            case DirectoryNotFoundException:
                return new UnauthorizedAccessException(
                    ErrorMessageBuilder.OutputDirectoryNotWritable(contextBasename));

            default:
                return new InvalidOperationException(ErrorMessageBuilder.ProcessingFailed());
        }
    }

    /// <summary>
    ///     Produces a sanitized message string describing an image-access failure. Used in
    ///     the response-field pattern (e.g. <c>PdfImageInfo.Error</c>) where an exception
    ///     cannot be thrown — the caller embeds the return value in the MCP response item
    ///     instead. Returns a fixed sentinel so that raw Aspose error text never surfaces to
    ///     the MCP caller.
    /// </summary>
    /// <param name="ex">
    ///     The raw exception from the image-access attempt. Must not be <c>null</c>. The
    ///     exception type is inspected but its <see cref="Exception.Message" /> is never
    ///     forwarded.
    /// </param>
    /// <returns>
    ///     A fixed sanitized error string: either the invalid-password sentinel (when
    ///     <paramref name="ex" /> is <see cref="Aspose.Pdf.InvalidPasswordException" />) or
    ///     <see cref="PdfErrorMessageBuilder.ImageAccessError" /> for all other cases.
    /// </returns>
    public static string TranslateImageAccessError(Exception ex)
    {
        if (ex is InvalidPasswordException)
            return ErrorMessageBuilder.InvalidPassword();

        return PdfErrorMessageBuilder.ImageAccessError();
    }
}
