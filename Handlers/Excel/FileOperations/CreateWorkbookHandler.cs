using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.FileOperations;

/// <summary>
///     Handler for creating new Excel workbooks.
/// </summary>
public class CreateWorkbookHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new Excel workbook.
    /// </summary>
    /// <param name="context">The workbook context (not used for create operation).</param>
    /// <param name="parameters">
    ///     Required: path or outputPath
    ///     Optional: sheetName
    /// </param>
    /// <returns>Success message with output path.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractCreateParameters(parameters);

        var targetPath = p.Path ?? p.OutputPath ?? throw new ArgumentException("path or outputPath is required");
        SecurityHelper.ValidateFilePath(targetPath, allowAbsolutePaths: true);

        using var workbook = new Workbook();

        if (!string.IsNullOrEmpty(p.SheetName))
            workbook.Worksheets[0].Name = p.SheetName;

        workbook.Save(targetPath);

        return Success($"Excel workbook created successfully. Output: {targetPath}");
    }

    /// <summary>
    ///     Extracts create parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted create parameters.</returns>
    private static CreateParameters ExtractCreateParameters(OperationParameters parameters)
    {
        return new CreateParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("sheetName"));
    }

    /// <summary>
    ///     Parameters for the create workbook operation.
    /// </summary>
    /// <param name="Path">The output file path.</param>
    /// <param name="OutputPath">Alternative output file path parameter.</param>
    /// <param name="SheetName">The name for the first worksheet.</param>
    private record CreateParameters(
        string? Path,
        string? OutputPath,
        string? SheetName);
}
