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
        var path = parameters.GetOptional<string?>("path");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var sheetName = parameters.GetOptional<string?>("sheetName");

        var targetPath = path ?? outputPath ?? throw new ArgumentException("path or outputPath is required");
        SecurityHelper.ValidateFilePath(targetPath, allowAbsolutePaths: true);

        using var workbook = new Workbook();

        if (!string.IsNullOrEmpty(sheetName))
            workbook.Worksheets[0].Name = sheetName;

        workbook.Save(targetPath);

        return Success($"Excel workbook created successfully. Output: {targetPath}");
    }
}
