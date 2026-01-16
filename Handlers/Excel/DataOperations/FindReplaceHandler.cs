using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for finding and replacing text in Excel worksheets.
/// </summary>
public class FindReplaceHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "find_replace";

    /// <summary>
    ///     Finds and replaces text in the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: findText, replaceText
    ///     Optional: sheetIndex, matchCase, matchEntireCell
    /// </param>
    /// <returns>Success message with replacement count.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var findReplaceParams = ExtractFindReplaceParameters(parameters);

        if (string.IsNullOrEmpty(findReplaceParams.FindText))
            throw new ArgumentException("findText is required for find_replace operation");
        if (findReplaceParams.ReplaceText == null)
            throw new ArgumentException("replaceText is required for find_replace operation");

        try
        {
            var workbook = context.Document;
            var totalReplacements = 0;
            var lookAt = findReplaceParams.MatchEntireCell ? LookAtType.EntireContent : LookAtType.Contains;

            if (findReplaceParams.SheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, findReplaceParams.SheetIndex.Value);
                totalReplacements += ReplaceInWorksheet(worksheet, findReplaceParams.FindText,
                    findReplaceParams.ReplaceText, findReplaceParams.MatchCase, lookAt);
            }
            else
            {
                foreach (var worksheet in workbook.Worksheets)
                    totalReplacements += ReplaceInWorksheet(worksheet, findReplaceParams.FindText,
                        findReplaceParams.ReplaceText, findReplaceParams.MatchCase, lookAt);
            }

            MarkModified(context);

            return Success(
                $"Replaced '{findReplaceParams.FindText}' with '{findReplaceParams.ReplaceText}' ({totalReplacements} replacements).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Extracts find and replace parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted find and replace parameters.</returns>
    private static FindReplaceParameters ExtractFindReplaceParameters(OperationParameters parameters)
    {
        return new FindReplaceParameters(
            parameters.GetOptional<string?>("findText"),
            parameters.GetOptional<string?>("replaceText"),
            parameters.GetOptional<int?>("sheetIndex"),
            parameters.GetOptional("matchCase", false),
            parameters.GetOptional("matchEntireCell", false)
        );
    }

    /// <summary>
    ///     Replaces text in a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to search in.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="matchCase">Whether to match case.</param>
    /// <param name="lookAt">The type of match (entire content or contains).</param>
    /// <returns>The number of replacements made.</returns>
    private static int ReplaceInWorksheet(Worksheet worksheet, string findText, string replaceText, bool matchCase,
        LookAtType lookAt)
    {
        var findOptions = new FindOptions
        {
            CaseSensitive = matchCase,
            LookAtType = lookAt
        };

        var replacedCells = new HashSet<string>();
        var cell = worksheet.Cells.Find(findText, null, findOptions);
        var count = 0;

        while (cell != null)
        {
            var cellName = cell.Name;
            if (replacedCells.Contains(cellName))
                break;

            if (lookAt == LookAtType.EntireContent)
            {
                cell.PutValue(replaceText);
            }
            else
            {
                var currentValue = cell.StringValue ?? "";
                var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                var newValue = currentValue.Replace(findText, replaceText, comparison);
                cell.PutValue(newValue);
            }

            replacedCells.Add(cellName);
            count++;
            cell = worksheet.Cells.Find(findText, cell, findOptions);
        }

        return count;
    }

    /// <summary>
    ///     Parameters for find and replace operation.
    /// </summary>
    /// <param name="FindText">The text to find.</param>
    /// <param name="ReplaceText">The text to replace with.</param>
    /// <param name="SheetIndex">The worksheet index (0-based), or null to search all sheets.</param>
    /// <param name="MatchCase">Whether to match case.</param>
    /// <param name="MatchEntireCell">Whether to match entire cell content.</param>
    private sealed record FindReplaceParameters(
        string? FindText,
        string? ReplaceText,
        int? SheetIndex,
        bool MatchCase,
        bool MatchEntireCell);
}
