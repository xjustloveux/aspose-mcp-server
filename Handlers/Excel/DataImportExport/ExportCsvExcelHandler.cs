using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Excel;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataImportExport;

namespace AsposeMcpServer.Handlers.Excel.DataImportExport;

/// <summary>
///     Handler for exporting Excel worksheet data to CSV format.
/// </summary>
[ResultType(typeof(ExportExcelResult))]
public class ExportCsvExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "export_csv";

    /// <summary>
    ///     Exports worksheet data to a CSV file.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputPath
    ///     Optional: sheetIndex (default: 0), separator (default: ',')
    /// </param>
    /// <returns>Export result with output path.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var separator = parameters.GetOptional("separator", ",");
        var sanitizeFormulas = parameters.GetOptional("sanitizeFormulas", true);

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for export_csv operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        try
        {
            var workbook = context.Document;
            ExcelHelper.ValidateSheetIndex(sheetIndex, workbook);

            var saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = separator.Length > 0 ? separator[0] : ','
            };

            // H16: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
            outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(outputPath));

            // Exporting is a read operation: the active sheet is a persisted workbook property, so
            // the temporary switch used to select the exported sheet is reverted before returning.
            var previousActiveSheet = workbook.Worksheets.ActiveSheetIndex;
            try
            {
                workbook.Worksheets.ActiveSheetIndex = sheetIndex;
                WriteCsv(workbook, outputPath, saveOptions, sanitizeFormulas,
                    saveOptions.Separator, context.ServerConfig?.AllowedBasePaths ?? [],
                    context.Recovery);
            }
            finally
            {
                workbook.Worksheets.ActiveSheetIndex = previousActiveSheet;
            }

            var worksheet = workbook.Worksheets[sheetIndex];
            var rowCount = worksheet.Cells.MaxDataRow + 1;

            return new ExportExcelResult
            {
                OutputPath = outputPath,
                RowCount = rowCount,
                Message = $"Sheet {sheetIndex} exported to CSV: {outputPath} ({rowCount} rows)."
            };
        }
        catch (CellsException ex)
        {
            throw CellsErrorTranslator.Translate(ex);
        }
    }

    /// <summary>
    ///     Writes the active worksheet as CSV, optionally neutralising formula injection first.
    /// </summary>
    /// <param name="workbook">The workbook whose active worksheet is exported.</param>
    /// <param name="outputPath">Resolved destination path.</param>
    /// <param name="saveOptions">Text save options carrying the configured separator.</param>
    /// <param name="sanitizeFormulas">
    ///     When <c>true</c> (default) every field that a spreadsheet application would evaluate is
    ///     prefixed with an apostrophe. Set to <c>false</c> for a byte-faithful export when the
    ///     consumer is a machine parser rather than a spreadsheet application.
    /// </param>
    /// <param name="separator">The separator character used in the produced text.</param>
    /// <param name="allowedBasePaths">Allowlist every path this publish touches is re-checked against.</param>
    /// <param name="recovery">
    ///     Where this host's publish records live and the key they are signed with (R18-ARCH01).
    /// </param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the export passes the output-size limit. It is thrown at the write that
    ///     crosses it, and the destination is left untouched.
    /// </exception>
    private static void WriteCsv(Workbook workbook, string outputPath, TxtSaveOptions saveOptions,
        bool sanitizeFormulas, char separator, IReadOnlyList<string> allowedBasePaths,
        RecoveryContext recovery)
    {
        // Both branches write somewhere else first and publish afterwards. Saving straight onto
        // the destination had no bound at all, and the sanitising branch measured its buffer only
        // once the whole sheet was already in memory — a size checked after the bytes exist is
        // not a limit, and a refusal that leaves a partial file behind is worse than none
        // (R3-R06). The staging and the publish are the shared transactional publisher's, not a
        // second copy of it with a short suffix and no containment re-check (R7-F01).
        BoundedFilePublisher.Publish(outputPath, RenderBudget.MaxOutputBytes, stream =>
        {
            if (sanitizeFormulas)
                WriteSanitised(workbook, stream, saveOptions, separator);
            else
                workbook.Save(stream, saveOptions);
        }, "CSV output", recovery, allowedBasePaths);
    }

    /// <summary>
    ///     Writes the worksheet with formula-injection neutralised.
    /// </summary>
    /// <param name="workbook">The workbook to export.</param>
    /// <param name="destination">
    ///     The publisher's stream. It is already bounded and already staged somewhere other than
    ///     the caller's file, so this no longer opens one of its own (R7-F01).
    /// </param>
    /// <param name="saveOptions">Text save options carrying the configured separator.</param>
    /// <param name="separator">The separator character used in the produced text.</param>
    /// <exception cref="ArgumentException">Thrown when the export passes the limit.</exception>
    private static void WriteSanitised(Workbook workbook, Stream destination,
        TxtSaveOptions saveOptions, char separator)
    {
        // Sanitising used to hold the sheet four times over at once — the save buffer, the byte
        // array it was copied into, the decoded string, and the sanitiser's output — so an export
        // inside the output limit could still need several times that limit in memory (R4-R03).
        // The export is held once and streamed through the sanitiser into the staging file.
        using var buffer = new MemoryStream();
        using (var bounded = new BoundedWriteStream(buffer, RenderBudget.MaxOutputBytes, "CSV output"))
        {
            workbook.Save(bounded, saveOptions);
        }

        buffer.Position = 0;

        // Both ends are measured in the bytes actually written rather than in UTF-16 code units,
        // which for CJK text are neither the same count nor the same size: the string measure
        // reported two bytes per character where three land on disk (R4-R03).
        using var reader = new StreamReader(buffer, new UTF8Encoding(false), false, 8192, true);
        using var writer = new StreamWriter(destination, new UTF8Encoding(false), 8192, true);

        CsvFormulaSanitizer.Sanitize(reader, writer, separator);
    }
}
