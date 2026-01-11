using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Filter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Filter;

public class ApplyExcelFilterHandlerTests : ExcelHandlerTestBase
{
    private readonly ApplyExcelFilterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Apply()
    {
        Assert.Equal("apply", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_AppliesFilterToCorrectSheet()
    {
        var workbook = CreateWorkbookWithData();
        workbook.Worksheets.Add("Sheet2");
        FillData(workbook.Worksheets[1]);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "range", "A1:B5" }
        });

        _handler.Execute(context, parameters);

        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
        Assert.Equal("A1:B5", workbook.Worksheets[1].AutoFilter.Range);
    }

    #endregion

    #region Basic Apply Operations

    [Fact]
    public void Execute_AppliesAutoFilter()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C10" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Auto filter applied", result);
        Assert.Contains("A1:C10", result);
        Assert.Equal("A1:C10", workbook.Worksheets[0].AutoFilter.Range);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSheetIndex()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 0", result);
    }

    [Theory]
    [InlineData("A1:A10")]
    [InlineData("B1:D20")]
    [InlineData("A1:Z100")]
    public void Execute_WithVariousRanges_AppliesFilter(string range)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", range }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(range, workbook.Worksheets[0].AutoFilter.Range);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "range", "A1:A10" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        FillData(workbook.Worksheets[0]);
        return workbook;
    }

    private static void FillData(Worksheet sheet)
    {
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["A2"].Value = "Item1";
        sheet.Cells["B2"].Value = 10;
        sheet.Cells["A3"].Value = "Item2";
        sheet.Cells["B3"].Value = 20;
    }

    #endregion
}
