using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Table;

/// <summary>
///     Handler for setting the style of a table (ListObject) in an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetStyleExcelTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_style";

    /// <summary>
    ///     Sets the style of a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: tableIndex, styleName
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

            GetExcelTablesHandler.ValidateTableIndex(worksheet, p.TableIndex);

            var listObject = worksheet.ListObjects[p.TableIndex];
            var tableStyleType = ResolveTableStyleType(p.StyleName);
            listObject.TableStyleType = tableStyleType;

            MarkModified(context);

            var tableName = listObject.DisplayName ?? $"Table{p.TableIndex + 1}";
            return new SuccessResult
            {
                Message = $"Table '{tableName}' style set to '{p.StyleName}'."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to set table style: {ex.Message}");
        }
    }

    /// <summary>
    ///     Resolves a style name string to a TableStyleType enum value.
    /// </summary>
    /// <param name="styleName">The style name (e.g., "TableStyleLight1", "TableStyleMedium9", "TableStyleDark1").</param>
    /// <returns>The corresponding TableStyleType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the style name is unknown.</exception>
    internal static TableStyleType ResolveTableStyleType(string styleName)
    {
        if (Enum.TryParse<TableStyleType>(styleName, true, out var result))
            return result;

        var availableStyles = string.Join(", ",
            Enum.GetNames<TableStyleType>().Take(10));
        throw new ArgumentException(
            $"Unknown table style: '{styleName}'. Available styles include: {availableStyles}... " +
            "Use TableStyleLight1-21, TableStyleMedium1-28, or TableStyleDark1-11.");
    }

    private static SetStyleParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var tableIndex = parameters.GetOptional<int?>("tableIndex");
        var styleName = parameters.GetOptional<string?>("styleName");

        if (!tableIndex.HasValue)
            throw new ArgumentException("tableIndex is required for set_style operation");
        if (string.IsNullOrEmpty(styleName))
            throw new ArgumentException("styleName is required for set_style operation");

        return new SetStyleParameters(sheetIndex, tableIndex.Value, styleName);
    }

    private sealed record SetStyleParameters(int SheetIndex, int TableIndex, string StyleName);
}
