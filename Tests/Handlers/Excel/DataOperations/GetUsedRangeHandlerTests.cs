using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

public class GetUsedRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly GetUsedRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetUsedRange()
    {
        Assert.Equal("get_used_range", _handler.Operation);
    }

    #endregion

    #region Basic Get Used Range Operations

    [Fact]
    public void Execute_ReturnsUsedRangeInfo()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Start");
        workbook.Worksheets[0].Cells["C3"].PutValue("End");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("worksheetName", result);
        Assert.Contains("firstRow", result);
        Assert.Contains("lastRow", result);
        Assert.Contains("firstColumn", result);
        Assert.Contains("lastColumn", result);
        Assert.Contains("range", result);
    }

    [Fact]
    public void Execute_WithSheetIndex_ReturnsSpecificSheetRange()
    {
        var workbook = CreateWorkbookWithSheets(2);
        workbook.Worksheets[1].Cells["B2"].PutValue("Data");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"sheetIndex\": 1", result);
    }

    [Fact]
    public void Execute_WithEmptySheet_ReturnsRangeInfo()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("worksheetName", result);
        Assert.Contains("sheetIndex", result);
    }

    [Fact]
    public void Execute_WithSingleCell_ReturnsCorrectRange()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["B2"].PutValue("Single");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B2:B2", result);
    }

    #endregion
}
