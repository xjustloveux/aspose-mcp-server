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
        var p = ExtractConvertParameters(parameters);

        if (string.IsNullOrEmpty(p.InputPath) && string.IsNullOrEmpty(p.SessionId))
            throw new ArgumentException("Either inputPath or sessionId is required for convert operation");
        if (string.IsNullOrEmpty(p.OutputPath))
            throw new ArgumentException("outputPath is required for convert operation");
        if (string.IsNullOrEmpty(p.Format))
            throw new ArgumentException("format is required for convert operation");

        Workbook workbook;
        string sourceDescription;

        if (!string.IsNullOrEmpty(p.SessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            workbook = context.SessionManager.GetDocument<Workbook>(p.SessionId, identity);
            sourceDescription = $"session {p.SessionId}";
        }
        else
        {
            workbook = new Workbook(p.InputPath);
            sourceDescription = p.InputPath!;
        }

        var saveFormat = p.Format.ToLower() switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "csv" => SaveFormat.Csv,
            "xlsx" => SaveFormat.Xlsx,
            "xls" => SaveFormat.Excel97To2003,
            "ods" => SaveFormat.Ods,
            "txt" => SaveFormat.TabDelimited,
            "tsv" => SaveFormat.TabDelimited,
            _ => throw new ArgumentException($"Unsupported format: {p.Format}")
        };

        workbook.Save(p.OutputPath, saveFormat);

        return Success($"Workbook from {sourceDescription} converted to {p.Format} format. Output: {p.OutputPath}");
    }

    /// <summary>
    ///     Extracts convert parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted convert parameters.</returns>
    private static ConvertParameters ExtractConvertParameters(OperationParameters parameters)
    {
        return new ConvertParameters(
            parameters.GetOptional<string?>("inputPath"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("format"));
    }

    /// <summary>
    ///     Parameters for the convert workbook operation.
    /// </summary>
    /// <param name="InputPath">The input file path.</param>
    /// <param name="SessionId">The session ID for session-based operations.</param>
    /// <param name="OutputPath">The output file path.</param>
    /// <param name="Format">The target format.</param>
    private sealed record ConvertParameters(
        string? InputPath,
        string? SessionId,
        string? OutputPath,
        string? Format);
}
