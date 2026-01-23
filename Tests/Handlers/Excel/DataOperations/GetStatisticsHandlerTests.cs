using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Results.Excel.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.True(result.TotalWorksheets > 0);
        Assert.NotNull(result.FileFormat);
        Assert.NotNull(result.Worksheets);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.Contains(result.Worksheets, w => w.Name == "Sheet2");
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        var worksheet = result.Worksheets[0];
        Assert.NotNull(worksheet.RangeStatistics);
        Assert.NotNull(worksheet.RangeStatistics!.Sum);
        Assert.NotNull(worksheet.RangeStatistics!.Average);
        Assert.NotNull(worksheet.RangeStatistics!.Min);
        Assert.NotNull(worksheet.RangeStatistics!.Max);
    }

    [Fact]
    public void Execute_IncludesChartAndPictureCount()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Data");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        var worksheet = result.Worksheets[0];
        Assert.True(worksheet.ChartsCount >= 0);
        Assert.True(worksheet.PicturesCount >= 0);
        Assert.True(worksheet.HyperlinksCount >= 0);
        Assert.True(worksheet.CommentsCount >= 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        var rangeStats = result.Worksheets[0].RangeStatistics;
        Assert.NotNull(rangeStats);
        Assert.True(rangeStats.NumericCells >= 0);
        Assert.True(rangeStats.NonNumericCells >= 0);
        Assert.True(rangeStats.EmptyCells >= 0);
    }

    #endregion
}
