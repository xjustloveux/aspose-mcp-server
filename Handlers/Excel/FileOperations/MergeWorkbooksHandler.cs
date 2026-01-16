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
        var p = ExtractMergeParameters(parameters);

        var targetPath = p.Path ?? p.OutputPath ?? throw new ArgumentException("path or outputPath is required");

        if (p.InputPaths == null || p.InputPaths.Length == 0)
            throw new ArgumentException("At least one input path is required");

        var validPaths = p.InputPaths.Where(path => !string.IsNullOrEmpty(path)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("No valid input paths provided");

        foreach (var inputPath in validPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);

        SecurityHelper.ValidateFilePath(targetPath, "outputPath", true);

        using var targetWorkbook = new Workbook(validPaths[0]);

        for (var i = 1; i < validPaths.Count; i++)
        {
            using var sourceWorkbook = new Workbook(validPaths[i]);

            if (p.MergeSheets)
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

    /// <summary>
    ///     Extracts merge parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted merge parameters.</returns>
    private static MergeParameters ExtractMergeParameters(OperationParameters parameters)
    {
        return new MergeParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string[]?>("inputPaths"),
            parameters.GetOptional("mergeSheets", false));
    }

    /// <summary>
    ///     Parameters for the merge workbooks operation.
    /// </summary>
    /// <param name="Path">The output file path.</param>
    /// <param name="OutputPath">Alternative output file path parameter.</param>
    /// <param name="InputPaths">The array of input file paths to merge.</param>
    /// <param name="MergeSheets">Whether to merge sheets with the same name.</param>
    private sealed record MergeParameters(
        string? Path,
        string? OutputPath,
        string[]? InputPaths,
        bool MergeSheets);
}
