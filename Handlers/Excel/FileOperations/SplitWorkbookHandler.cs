using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Handlers.Excel.FileOperations;

/// <summary>
///     Handler for splitting an Excel workbook into separate files per worksheet.
/// </summary>
public class SplitWorkbookHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "split";

    /// <summary>
    ///     Splits an Excel workbook into separate files, one per worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: outputDirectory
    ///     Optional: inputPath, path, sessionId (one is required), sheetIndices, outputFileNamePattern
    /// </param>
    /// <returns>Success message with split details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var inputPath = parameters.GetOptional<string?>("inputPath");
        var path = parameters.GetOptional<string?>("path");
        var sessionId = parameters.GetOptional<string?>("sessionId");
        var outputDirectory = parameters.GetOptional<string?>("outputDirectory");
        var sheetIndices = parameters.GetOptional<int[]?>("sheetIndices");
        var outputFileNamePattern = parameters.GetOptional("outputFileNamePattern", "sheet_{name}.xlsx");

        var sourcePath = inputPath ?? path;
        if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath, path, or sessionId is required");
        if (string.IsNullOrEmpty(outputDirectory))
            throw new ArgumentException("outputDirectory is required for split operation");

        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        Workbook sourceWorkbook;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            sourceWorkbook = context.SessionManager.GetDocument<Workbook>(sessionId, identity);
        }
        else
        {
            sourceWorkbook = new Workbook(sourcePath);
        }

        var indicesToSplit = sheetIndices is { Length: > 0 }
            ? sheetIndices.Distinct().ToList()
            : Enumerable.Range(0, sourceWorkbook.Worksheets.Count).ToList();

        List<string> splitFiles = [];

        foreach (var sheetIndex in indicesToSplit)
        {
            if (sheetIndex < 0 || sheetIndex >= sourceWorkbook.Worksheets.Count)
                continue;

            var worksheet = sourceWorkbook.Worksheets[sheetIndex];
            var fileName = outputFileNamePattern
                .Replace("{index}", sheetIndex.ToString())
                .Replace("{name}", worksheet.Name);
            var outputFilePath = Path.Combine(outputDirectory, fileName);

            using var newWorkbook = new Workbook();
            newWorkbook.Worksheets.RemoveAt(0);
            newWorkbook.Worksheets.Add(worksheet.Name);
            var newSheet = newWorkbook.Worksheets[0];
            var maxRow = worksheet.Cells.MaxDataRow;
            var maxCol = worksheet.Cells.MaxDataColumn;

            if (maxRow >= 0 && maxCol >= 0)
            {
                var sourceRange = worksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                var destRange = newSheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
            }

            newWorkbook.Save(outputFilePath);
            splitFiles.Add(outputFilePath);
        }

        return Success($"Split workbook into {splitFiles.Count} files. Output: {outputDirectory}");
    }
}
