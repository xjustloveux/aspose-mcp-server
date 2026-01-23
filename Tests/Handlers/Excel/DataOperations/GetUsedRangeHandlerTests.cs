using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Results.Excel.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetUsedRangeResult>(res);

        Assert.NotNull(result.WorksheetName);
        Assert.True(result.FirstRow >= 0);
        Assert.True(result.LastRow >= 0);
        Assert.True(result.FirstColumn >= 0);
        Assert.True(result.LastColumn >= 0);
        Assert.NotNull(result.Range);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetUsedRangeResult>(res);

        Assert.Equal(1, result.SheetIndex);
    }

    [Fact]
    public void Execute_WithEmptySheet_ReturnsRangeInfo()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetUsedRangeResult>(res);

        Assert.NotNull(result.WorksheetName);
        Assert.Equal(0, result.SheetIndex);
    }

    [Fact]
    public void Execute_WithSingleCell_ReturnsCorrectRange()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["B2"].PutValue("Single");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetUsedRangeResult>(res);

        Assert.Equal("B2:B2", result.Range);
    }

    #endregion
}
