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
        var p = ExtractSplitParameters(parameters);

        ValidateParameters(p.InputPath, p.Path, p.SessionId, p.OutputDirectory);
        EnsureOutputDirectoryExists(p.OutputDirectory!);

        var sourceWorkbook = GetSourceWorkbook(context, p.SessionId, p.InputPath ?? p.Path);
        var indicesToSplit = GetIndicesToSplit(p.SheetIndices, sourceWorkbook.Worksheets.Count);
        var splitFiles = SplitWorksheets(sourceWorkbook, indicesToSplit, p.OutputDirectory!, p.OutputFileNamePattern);

        return Success($"Split workbook into {splitFiles.Count} files. Output: {p.OutputDirectory}");
    }

    /// <summary>
    ///     Validates the required parameters for the split operation.
    /// </summary>
    /// <param name="inputPath">The input file path.</param>
    /// <param name="path">Alternative path parameter.</param>
    /// <param name="sessionId">The session ID for session-based operations.</param>
    /// <param name="outputDirectory">The output directory path.</param>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    private static void ValidateParameters(string? inputPath, string? path, string? sessionId, string? outputDirectory)
    {
        var sourcePath = inputPath ?? path;
        if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath, path, or sessionId is required");
        if (string.IsNullOrEmpty(outputDirectory))
            throw new ArgumentException("outputDirectory is required for split operation");
    }

    /// <summary>
    ///     Ensures the output directory exists, creating it if necessary.
    /// </summary>
    /// <param name="outputDirectory">The output directory path.</param>
    private static void EnsureOutputDirectoryExists(string outputDirectory)
    {
        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);
    }

    /// <summary>
    ///     Gets the source workbook from file path or session.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="sessionId">The session ID if using session-based workbook.</param>
    /// <param name="sourcePath">The file path if loading from disk.</param>
    /// <returns>The source workbook.</returns>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    private static Workbook GetSourceWorkbook(OperationContext<Workbook> context, string? sessionId, string? sourcePath)
    {
        if (string.IsNullOrEmpty(sessionId))
            return new Workbook(sourcePath);

        if (context.SessionManager == null)
            throw new InvalidOperationException("Session management is not enabled");

        var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
        return context.SessionManager.GetDocument<Workbook>(sessionId, identity);
    }

    /// <summary>
    ///     Gets the list of worksheet indices to split.
    /// </summary>
    /// <param name="sheetIndices">The specified sheet indices, or null for all sheets.</param>
    /// <param name="worksheetCount">The total number of worksheets.</param>
    /// <returns>A list of worksheet indices to process.</returns>
    private static List<int> GetIndicesToSplit(int[]? sheetIndices, int worksheetCount)
    {
        return sheetIndices is { Length: > 0 }
            ? sheetIndices.Distinct().ToList()
            : Enumerable.Range(0, worksheetCount).ToList();
    }

    /// <summary>
    ///     Splits worksheets from the source workbook into separate files.
    /// </summary>
    /// <param name="sourceWorkbook">The source workbook to split.</param>
    /// <param name="indicesToSplit">The list of worksheet indices to split.</param>
    /// <param name="outputDirectory">The output directory path.</param>
    /// <param name="outputFileNamePattern">The output file name pattern.</param>
    /// <returns>A list of created file paths.</returns>
    private static List<string> SplitWorksheets(Workbook sourceWorkbook, List<int> indicesToSplit,
        string outputDirectory, string outputFileNamePattern)
    {
        List<string> splitFiles = [];

        foreach (var sheetIndex in indicesToSplit)
        {
            if (sheetIndex < 0 || sheetIndex >= sourceWorkbook.Worksheets.Count)
                continue;

            var outputFilePath =
                SaveWorksheetAsNewFile(sourceWorkbook, sheetIndex, outputDirectory, outputFileNamePattern);
            splitFiles.Add(outputFilePath);
        }

        return splitFiles;
    }

    /// <summary>
    ///     Saves a single worksheet as a new workbook file.
    /// </summary>
    /// <param name="sourceWorkbook">The source workbook.</param>
    /// <param name="sheetIndex">The index of the worksheet to save.</param>
    /// <param name="outputDirectory">The output directory path.</param>
    /// <param name="outputFileNamePattern">The output file name pattern with placeholders.</param>
    /// <returns>The path to the created file.</returns>
    private static string SaveWorksheetAsNewFile(Workbook sourceWorkbook, int sheetIndex,
        string outputDirectory, string outputFileNamePattern)
    {
        var worksheet = sourceWorkbook.Worksheets[sheetIndex];
        var fileName = outputFileNamePattern
            .Replace("{index}", sheetIndex.ToString())
            .Replace("{name}", worksheet.Name);
        var outputFilePath = Path.Combine(outputDirectory, fileName);

        using var newWorkbook = new Workbook();
        newWorkbook.Worksheets.RemoveAt(0);
        newWorkbook.Worksheets.Add(worksheet.Name);
        var newSheet = newWorkbook.Worksheets[0];

        CopyWorksheetData(worksheet, newSheet);
        newWorkbook.Save(outputFilePath);

        return outputFilePath;
    }

    /// <summary>
    ///     Copies all data from a source worksheet to a destination worksheet.
    /// </summary>
    /// <param name="sourceWorksheet">The source worksheet to copy from.</param>
    /// <param name="destWorksheet">The destination worksheet to copy to.</param>
    private static void CopyWorksheetData(Worksheet sourceWorksheet, Worksheet destWorksheet)
    {
        var maxRow = sourceWorksheet.Cells.MaxDataRow;
        var maxCol = sourceWorksheet.Cells.MaxDataColumn;

        if (maxRow < 0 || maxCol < 0) return;

        var sourceRange = sourceWorksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
        var destRange = destWorksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
        destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
    }

    /// <summary>
    ///     Extracts split parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted split parameters.</returns>
    private static SplitParameters ExtractSplitParameters(OperationParameters parameters)
    {
        return new SplitParameters(
            parameters.GetOptional<string?>("inputPath"),
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetOptional<string?>("outputDirectory"),
            parameters.GetOptional<int[]?>("sheetIndices"),
            parameters.GetOptional("outputFileNamePattern", "sheet_{name}.xlsx"));
    }

    /// <summary>
    ///     Parameters for the split workbook operation.
    /// </summary>
    /// <param name="InputPath">The input file path.</param>
    /// <param name="Path">Alternative path parameter.</param>
    /// <param name="SessionId">The session ID for session-based operations.</param>
    /// <param name="OutputDirectory">The output directory path.</param>
    /// <param name="SheetIndices">The indices of sheets to split.</param>
    /// <param name="OutputFileNamePattern">The output file name pattern.</param>
    private sealed record SplitParameters(
        string? InputPath,
        string? Path,
        string? SessionId,
        string? OutputDirectory,
        int[]? SheetIndices,
        string OutputFileNamePattern);
}
