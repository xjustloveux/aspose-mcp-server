using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.FileOperations;

/// <summary>
///     Handler for merging multiple Excel workbooks into one.
/// </summary>
public class MergeWorkbooksHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "merge";

    /// <summary>
    ///     Merges multiple Excel workbooks into one.
    /// </summary>
    /// <param name="context">The workbook context (not used for merge operation).</param>
    /// <param name="parameters">
    ///     Required: path or outputPath, inputPaths
    ///     Optional: mergeSheets (default: false)
    /// </param>
    /// <returns>Success message with merge details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var path = parameters.GetOptional<string?>("path");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var inputPaths = parameters.GetOptional<string[]?>("inputPaths");
        var mergeSheets = parameters.GetOptional("mergeSheets", false);

        var targetPath = path ?? outputPath ?? throw new ArgumentException("path or outputPath is required");

        if (inputPaths == null || inputPaths.Length == 0)
            throw new ArgumentException("At least one input path is required");

        var validPaths = inputPaths.Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("No valid input paths provided");

        foreach (var inputPath in validPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);

        SecurityHelper.ValidateFilePath(targetPath, "outputPath", true);

        using var targetWorkbook = new Workbook(validPaths[0]);

        for (var i = 1; i < validPaths.Count; i++)
        {
            using var sourceWorkbook = new Workbook(validPaths[i]);

            if (mergeSheets)
                MergeSheetsWithSameName(targetWorkbook, sourceWorkbook);
            else
                targetWorkbook.Combine(sourceWorkbook);
        }

        targetWorkbook.Save(targetPath);

        return Success($"Merged {validPaths.Count} workbooks successfully. Output: {targetPath}");
    }

    /// <summary>
    ///     Merges worksheets from source workbook into target workbook,
    ///     appending data to sheets with matching names or creating new sheets.
    /// </summary>
    /// <param name="targetWorkbook">The target workbook to merge into.</param>
    /// <param name="sourceWorkbook">The source workbook to merge from.</param>
    private static void MergeSheetsWithSameName(Workbook targetWorkbook, Workbook sourceWorkbook)
    {
        foreach (var sourceSheet in sourceWorkbook.Worksheets)
        {
            var existingSheet = targetWorkbook.Worksheets[sourceSheet.Name];
            if (existingSheet != null)
            {
                var lastRow = existingSheet.Cells.MaxDataRow + 1;
                var sourceMaxRow = sourceSheet.Cells.MaxDataRow;
                var sourceMaxCol = sourceSheet.Cells.MaxDataColumn;

                if (sourceMaxRow >= 0 && sourceMaxCol >= 0)
                {
                    var sourceRange = sourceSheet.Cells.CreateRange(0, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                    var destRange = existingSheet.Cells.CreateRange(lastRow, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                    destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                }
            }
            else
            {
                targetWorkbook.Worksheets.Add(sourceSheet.Name);
                var newSheet = targetWorkbook.Worksheets[^1];
                var sourceMaxRow = sourceSheet.Cells.MaxDataRow;
                var sourceMaxCol = sourceSheet.Cells.MaxDataColumn;

                if (sourceMaxRow >= 0 && sourceMaxCol >= 0)
                {
                    var sourceRange = sourceSheet.Cells.CreateRange(0, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                    var destRange = newSheet.Cells.CreateRange(0, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                    destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                }
            }
        }
    }
}
