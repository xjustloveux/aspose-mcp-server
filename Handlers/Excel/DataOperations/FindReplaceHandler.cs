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
        var findText = parameters.GetOptional<string?>("findText");
        var replaceText = parameters.GetOptional<string?>("replaceText");
        var sheetIndex = parameters.GetOptional<int?>("sheetIndex");
        var matchCase = parameters.GetOptional("matchCase", false);
        var matchEntireCell = parameters.GetOptional("matchEntireCell", false);

        if (string.IsNullOrEmpty(findText))
            throw new ArgumentException("findText is required for find_replace operation");
        if (replaceText == null)
            throw new ArgumentException("replaceText is required for find_replace operation");

        try
        {
            var workbook = context.Document;
            var totalReplacements = 0;
            var lookAt = matchEntireCell ? LookAtType.EntireContent : LookAtType.Contains;

            if (sheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
                totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
            }
            else
            {
                foreach (var worksheet in workbook.Worksheets)
                    totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
            }

            MarkModified(context);

            return Success($"Replaced '{findText}' with '{replaceText}' ({totalReplacements} replacements).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
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
}
