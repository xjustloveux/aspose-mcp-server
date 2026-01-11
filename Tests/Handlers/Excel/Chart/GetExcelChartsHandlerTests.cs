using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No charts found", json.RootElement.GetProperty("message").GetString());
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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("count", out _));
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateWorkbookWithCharts(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(1, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.NotNull(json.RootElement.GetProperty("worksheetName").GetString());
    }

    #endregion

    #region Chart Details

    [Fact]
    public void Execute_ReturnsChartIndex()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstChart = json.RootElement.GetProperty("items")[0];

        Assert.Equal(0, firstChart.GetProperty("index").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsChartType()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstChart = json.RootElement.GetProperty("items")[0];

        Assert.NotNull(firstChart.GetProperty("type").GetString());
    }

    [Fact]
    public void Execute_ReturnsLocation()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstChart = json.RootElement.GetProperty("items")[0];

        Assert.True(firstChart.TryGetProperty("location", out _));
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
