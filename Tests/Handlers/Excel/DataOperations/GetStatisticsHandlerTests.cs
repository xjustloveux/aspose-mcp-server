using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

public class GetStatisticsHandlerTests : ExcelHandlerTestBase
{
    private readonly GetStatisticsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStatistics()
    {
        Assert.Equal("get_statistics", _handler.Operation);
    }

    #endregion

    #region Basic Get Statistics Operations

    [Fact]
    public void Execute_ReturnsWorkbookStatistics()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Data");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("totalWorksheets", result);
        Assert.Contains("fileFormat", result);
        Assert.Contains("worksheets", result);
    }

    [Fact]
    public void Execute_WithSheetIndex_ReturnsSpecificSheetStatistics()
    {
        var workbook = CreateWorkbookWithSheets(2);
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets[1].Name = "Sheet2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Sheet2", result);
    }

    [Fact]
    public void Execute_WithRange_ReturnsRangeStatistics()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(10);
        workbook.Worksheets[0].Cells["A2"].PutValue(20);
        workbook.Worksheets[0].Cells["A3"].PutValue(30);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("rangeStatistics", result);
        Assert.Contains("sum", result);
        Assert.Contains("average", result);
        Assert.Contains("min", result);
        Assert.Contains("max", result);
    }

    [Fact]
    public void Execute_IncludesChartAndPictureCount()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Data");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("chartsCount", result);
        Assert.Contains("picturesCount", result);
        Assert.Contains("hyperlinksCount", result);
        Assert.Contains("commentsCount", result);
    }

    [Fact]
    public void Execute_WithMixedData_CountsCorrectly()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(10);
        workbook.Worksheets[0].Cells["A2"].PutValue("Text");
        workbook.Worksheets[0].Cells["A3"].PutValue(30);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A4" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("numericCells", result);
        Assert.Contains("nonNumericCells", result);
        Assert.Contains("emptyCells", result);
    }

    #endregion
}
