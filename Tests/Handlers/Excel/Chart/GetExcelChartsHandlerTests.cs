using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;
using AsposeMcpServer.Results.Excel.Chart;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Chart;

public class GetExcelChartsHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelChartsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region No Charts

    [Fact]
    public void Execute_NoCharts_ReturnsEmptyResult()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Contains("No charts found", result.Message);
    }

    #endregion

    #region Filter By Sheet

    [Fact]
    public void Execute_WithSheetIndex_GetsFromCorrectSheet()
    {
        var workbook = CreateWorkbookWithChart();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.Equal(0, result.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsCharts()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.True(result.Count >= 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateWorkbookWithCharts(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.Equal(2, result.Count);
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.Single(result.Items);
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.NotNull(result.WorksheetName);
    }

    #endregion

    #region Chart Details

    [Fact]
    public void Execute_ReturnsChartIndex()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.Equal(0, result.Items[0].Index);
    }

    [Fact]
    public void Execute_ReturnsChartType()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.NotNull(result.Items[0].Type);
    }

    [Fact]
    public void Execute_ReturnsLocation()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartsResult>(res);

        Assert.NotNull(result.Items[0].Location);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithChart()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["A2"].Value = 20;
        sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        sheet.Charts[0].NSeries.Add("A1:A2", true);
        return workbook;
    }

    private static Workbook CreateWorkbookWithCharts(int count)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["A2"].Value = 20;

        for (var i = 0; i < count; i++)
        {
            sheet.Charts.Add(ChartType.Column, 5 + i * 20, 0, 20 + i * 20, 8);
            sheet.Charts[i].NSeries.Add("A1:A2", true);
        }

        return workbook;
    }

    #endregion
}
