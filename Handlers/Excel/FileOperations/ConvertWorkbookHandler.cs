using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Handlers.Excel.FileOperations;

/// <summary>
///     Handler for converting Excel workbooks to different formats.
/// </summary>
public class ConvertWorkbookHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "convert";

    /// <summary>
    ///     Converts an Excel workbook to a different format.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: outputPath, format
    ///     Optional: inputPath, sessionId (one of inputPath or sessionId is required if not using context document)
    /// </param>
    /// <returns>Success message with conversion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var inputPath = parameters.GetOptional<string?>("inputPath");
        var sessionId = parameters.GetOptional<string?>("sessionId");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var format = parameters.GetOptional<string?>("format");

        if (string.IsNullOrEmpty(inputPath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath or sessionId is required for convert operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for convert operation");
        if (string.IsNullOrEmpty(format))
            throw new ArgumentException("format is required for convert operation");

        Workbook workbook;
        string sourceDescription;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            workbook = context.SessionManager.GetDocument<Workbook>(sessionId, identity);
            sourceDescription = $"session {sessionId}";
        }
        else
        {
            workbook = new Workbook(inputPath);
            sourceDescription = inputPath!;
        }

        var saveFormat = format.ToLower() switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "csv" => SaveFormat.Csv,
            "xlsx" => SaveFormat.Xlsx,
            "xls" => SaveFormat.Excel97To2003,
            "ods" => SaveFormat.Ods,
            "txt" => SaveFormat.TabDelimited,
            "tsv" => SaveFormat.TabDelimited,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        workbook.Save(outputPath, saveFormat);

        return Success($"Workbook from {sourceDescription} converted to {format} format. Output: {outputPath}");
    }
}
