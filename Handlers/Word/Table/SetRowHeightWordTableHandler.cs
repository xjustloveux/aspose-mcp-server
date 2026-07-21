using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for setting row height in Word document tables.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetRowHeightWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_row_height";

    /// <summary>
    ///     Sets row height for a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: rowIndex, rowHeight.
    ///     Optional: tableIndex (default 0), heightRule (default "atLeast"), sectionIndex.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetRowHeightParameters(parameters);

        var doc = context.Document;
        // Omitted sectionIndex means the document-wide flat table index, matching 'get'.
        var table = WordTableHelper.GetTable(doc, p.TableIndex, p.SectionIndex);
        if (p.RowIndex < 0 || p.RowIndex >= table.Rows.Count)
            throw new ArgumentException($"Row index {p.RowIndex} out of range");

        var row = table.Rows[p.RowIndex];
        var isAuto = p.HeightRule.Equals("auto", StringComparison.OrdinalIgnoreCase);
        row.RowFormat.HeightRule = p.HeightRule.ToLower() switch
        {
            "auto" => HeightRule.Auto,
            "atleast" => HeightRule.AtLeast,
            "exactly" => HeightRule.Exactly,
            _ => HeightRule.AtLeast
        };

        if (!isAuto)
            row.RowFormat.Height = p.RowHeight;

        MarkModified(context);

        return new SuccessResult
            { Message = $"Successfully set row {p.RowIndex} height to {p.RowHeight} pt ({p.HeightRule})." };
    }

    private static SetRowHeightParameters ExtractSetRowHeightParameters(OperationParameters parameters)
    {
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        var rowHeight = parameters.GetOptional<double?>("rowHeight");

        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for set_row_height operation");
        if (!rowHeight.HasValue)
            throw new ArgumentException("rowHeight is required for set_row_height operation");
        if (rowHeight.Value <= 0)
            throw new ArgumentException($"Row height {rowHeight.Value} must be greater than 0");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var heightRule = parameters.GetOptional("heightRule", "atLeast");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new SetRowHeightParameters(rowIndex.Value, rowHeight.Value, tableIndex, heightRule, sectionIndex);
    }

    private sealed record SetRowHeightParameters(
        int RowIndex,
        double RowHeight,
        int TableIndex,
        string HeightRule,
        int? SectionIndex);
}
