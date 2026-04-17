using Aspose.Cells;

namespace AsposeMcpServer.Errors.Excel;

/// <summary>
///     Translates raw Aspose.Cells / BCL exceptions into sanitized BCL exceptions for the
///     Excel tool surface. No inner-exception <see cref="Exception.Message" /> text is ever
///     forwarded to the emitted error — only fixed sentinels from
///     <see cref="ErrorMessageBuilder" /> flow through. Modelled after
///     <c>OleErrorTranslator</c> so that identical failure modes produce identical BCL types
///     and sanitized messages across the OLE and Excel tool families.
/// </summary>
public static class CellsErrorTranslator
{
    /// <summary>
    ///     Maps an arbitrary exception thrown during an Excel (Aspose.Cells) operation to a
    ///     sanitized BCL exception. The mapping is:
    ///     <list type="bullet">
    ///         <item><see cref="ExceptionType.IncorrectPassword" /> → <see cref="UnauthorizedAccessException" /></item>
    ///         <item>
    ///             <see cref="ExceptionType.FileFormat" /> or <see cref="ExceptionType.UnsupportedFeature" /> or
    ///             <see cref="ExceptionType.UnsupportedStream" /> → <see cref="NotSupportedException" />
    ///         </item>
    ///         <item>
    ///             <see cref="ExceptionType.IO" /> or <see cref="ExceptionType.Permission" /> →
    ///             <see cref="UnauthorizedAccessException" /> (output-directory context)
    ///         </item>
    ///         <item>
    ///             <see cref="ExceptionType.FileCorrupted" /> or <see cref="ExceptionType.InvalidData" /> →
    ///             <see cref="InvalidOperationException" />
    ///         </item>
    ///         <item><see cref="ExceptionType.Limitation" /> → <see cref="NotSupportedException" /></item>
    ///         <item>All other <see cref="CellsException" /> codes → <see cref="InvalidOperationException" /></item>
    ///         <item><see cref="ArgumentOutOfRangeException" /> → <see cref="ArgumentOutOfRangeException" /> (sanitized)</item>
    ///         <item>
    ///             <see cref="UnauthorizedAccessException" /> or <see cref="DirectoryNotFoundException" /> →
    ///             <see cref="UnauthorizedAccessException" /> (output-directory context)
    ///         </item>
    ///         <item>All other → <see cref="InvalidOperationException" /></item>
    ///     </list>
    /// </summary>
    /// <param name="ex">
    ///     The raw exception thrown from an Aspose.Cells API or IO call. Must not be
    ///     <c>null</c>.
    /// </param>
    /// <param name="contextBasename">
    ///     Optional sanitized basename of the file being processed (e.g. the source
    ///     filename). Used in output-directory messages to add context without leaking a
    ///     full path. May be <c>null</c>.
    /// </param>
    /// <returns>
    ///     A sanitized BCL exception ready to be thrown. The returned exception is always a
    ///     new instance — the original <paramref name="ex" /> is never re-thrown or attached
    ///     as an inner exception, so internal Aspose stack frames do not escape.
    /// </returns>
    public static Exception Translate(Exception ex, string? contextBasename = null)
    {
        if (ex is CellsException cellsEx)
            return cellsEx.Code switch
            {
                ExceptionType.IncorrectPassword =>
                    new UnauthorizedAccessException(ErrorMessageBuilder.InvalidPassword()),

                ExceptionType.FileFormat or
                    ExceptionType.UnsupportedFeature or
                    ExceptionType.UnsupportedStream or
                    ExceptionType.Limitation =>
                    new NotSupportedException(ErrorMessageBuilder.UnsupportedFormat()),

                ExceptionType.IO or
                    ExceptionType.Permission =>
                    new UnauthorizedAccessException(
                        ErrorMessageBuilder.OutputDirectoryNotWritable(contextBasename)),

                ExceptionType.FileCorrupted or
                    ExceptionType.InvalidData =>
                    new InvalidOperationException(ErrorMessageBuilder.ProcessingFailed()),

                // Covers Chart, DataType, DataValidation, ConditionalFormatting, Formula,
                // InvalidOperator, License, PivotTable, Shape, Sparkline, SheetName,
                // SheetType, Interrupted, PageSetup, UndisclosedInformation, and any future
                // codes not yet enumerated.
                _ => new InvalidOperationException(ErrorMessageBuilder.ProcessingFailed())
            };

        switch (ex)
        {
            case ArgumentOutOfRangeException:
                return new ArgumentOutOfRangeException(
                    nameof(ex), ErrorMessageBuilder.IndexOutOfRange());

            case UnauthorizedAccessException:
            case DirectoryNotFoundException:
                return new UnauthorizedAccessException(
                    ErrorMessageBuilder.OutputDirectoryNotWritable(contextBasename));

            default:
                return new InvalidOperationException(ErrorMessageBuilder.ProcessingFailed());
        }
    }
}
