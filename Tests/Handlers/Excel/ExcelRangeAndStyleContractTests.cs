using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Handlers.Excel.Style;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel;

/// <summary>
///     Guards RB-15 and RB-16. Moving a range copied it and then cleared the source
///     unconditionally, so a destination overlapping the source lost the data that had just been
///     written there. Formatting created a blank style and applied it with <c>StyleFlag.All</c>, so
///     a later call that set one attribute reset every attribute an earlier call had applied.
/// </summary>
public class ExcelRangeAndStyleContractTests : TestBase
{
    /// <summary>
    ///     Builds a workbook whose first column holds the supplied values from row 1 down.
    /// </summary>
    /// <param name="values">Values for A1, A2, ...</param>
    /// <returns>The workbook.</returns>
    private static Workbook WorkbookWithColumn(params string[] values)
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;
        for (var i = 0; i < values.Length; i++) cells[i, 0].PutValue(values[i]);
        return workbook;
    }

    /// <summary>Runs the move operation.</summary>
    /// <param name="workbook">Workbook to modify.</param>
    /// <param name="sourceRange">Source range in A1 notation.</param>
    /// <param name="destCell">Destination top-left cell.</param>
    private static void Move(Workbook workbook, string sourceRange, string destCell)
    {
        var handler = new MoveExcelRangeHandler();
        var context = new OperationContext<Workbook> { Document = workbook };
        var parameters = new OperationParameters();
        parameters.Set("sourceRange", sourceRange);
        parameters.Set("destCell", destCell);
        handler.Execute(context, parameters);
    }

    /// <summary>Runs the format operation with the supplied parameters.</summary>
    /// <param name="workbook">Workbook to modify.</param>
    /// <param name="values">Parameter names and values.</param>
    private static void Format(Workbook workbook, Dictionary<string, object?> values)
    {
        var handler = new FormatCellsHandler();
        var context = new OperationContext<Workbook> { Document = workbook };
        var parameters = new OperationParameters();
        foreach (var (key, value) in values) parameters.Set(key, value);
        handler.Execute(context, parameters);
    }

    [Fact]
    public void Move_DownwardOverlap_ShouldKeepEveryValue()
    {
        var workbook = WorkbookWithColumn("v1", "v2", "v3", "v4", "v5");

        Move(workbook, "A1:A5", "A2");

        var cells = workbook.Worksheets[0].Cells;
        Assert.Equal("v1", cells["A2"].StringValue);
        Assert.Equal("v5", cells["A6"].StringValue);
        Assert.Equal(string.Empty, cells["A1"].StringValue);
    }

    [Fact]
    public void Move_UpwardOverlap_ShouldKeepEveryValue()
    {
        var workbook = WorkbookWithColumn("v1", "v2", "v3", "v4", "v5");

        Move(workbook, "A2:A5", "A1");

        var cells = workbook.Worksheets[0].Cells;
        Assert.Equal("v2", cells["A1"].StringValue);
        Assert.Equal("v5", cells["A4"].StringValue);
    }

    [Fact]
    public void Move_WithoutOverlap_ShouldStillClearTheSource()
    {
        var workbook = WorkbookWithColumn("v1", "v2");

        Move(workbook, "A1:A2", "C1");

        var cells = workbook.Worksheets[0].Cells;
        Assert.Equal("v1", cells["C1"].StringValue);
        Assert.Equal("v2", cells["C2"].StringValue);
        Assert.Equal(string.Empty, cells["A1"].StringValue);
        Assert.Equal(string.Empty, cells["A2"].StringValue);
    }

    [Fact]
    public void Format_SecondCallSettingOnlyBold_ShouldKeepTheEarlierBackground()
    {
        var workbook = WorkbookWithColumn("data");

        Format(workbook, new Dictionary<string, object?>
        {
            ["range"] = "A1",
            ["backgroundColor"] = "#FF0000"
        });
        Format(workbook, new Dictionary<string, object?>
        {
            ["range"] = "A1",
            ["bold"] = true
        });

        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.True(style.Font.IsBold);
        Assert.Equal(255, style.ForegroundColor.R);
        Assert.Equal(0, style.ForegroundColor.G);
    }

    [Fact]
    public void Format_SecondCallSettingOnlyBackground_ShouldKeepTheEarlierNumberFormat()
    {
        var workbook = WorkbookWithColumn("1234");

        Format(workbook, new Dictionary<string, object?>
        {
            ["range"] = "A1",
            ["numberFormat"] = "0.00"
        });
        var applied = workbook.Worksheets[0].Cells["A1"].GetStyle().Custom;

        Format(workbook, new Dictionary<string, object?>
        {
            ["range"] = "A1",
            ["backgroundColor"] = "#00FF00"
        });

        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(applied, style.Custom);
        Assert.Equal(0, style.ForegroundColor.R);
        Assert.Equal(255, style.ForegroundColor.G);
    }

    [Fact]
    public void Format_SingleCall_ShouldStillApplyEverythingItWasGiven()
    {
        var workbook = WorkbookWithColumn("data");

        Format(workbook, new Dictionary<string, object?>
        {
            ["range"] = "A1",
            ["bold"] = true,
            ["italic"] = true,
            ["fontSize"] = 14,
            ["backgroundColor"] = "#0000FF",
            ["numberFormat"] = "0.00"
        });

        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.True(style.Font.IsBold);
        Assert.True(style.Font.IsItalic);
        Assert.Equal(14, style.Font.Size);
        Assert.Equal(255, style.ForegroundColor.B);
        Assert.Equal("0.00", style.Custom);
    }
}
