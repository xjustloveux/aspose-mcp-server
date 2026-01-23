using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Chart;

public class UpdateExcelChartDataHandlerTests : ExcelHandlerTestBase
{
    private readonly UpdateExcelChartDataHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_UpdateData()
    {
        Assert.Equal("update_data", _handler.Operation);
    }

    #endregion

    #region Category Axis

    [Fact]
    public void Execute_WithCategoryAxisDataRange_IncludesInResult()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "B1:B3" },
            { "categoryAxisDataRange", "A1:A3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("X-axis: A1:A3", result.Message);
    }

    #endregion

    #region Different Chart Index

    [Fact]
    public void Execute_WithChartIndex_UpdatesCorrectChart()
    {
        var workbook = CreateWorkbookWithMultipleCharts();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 1 },
            { "dataRange", "B1:B3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#1", result.Message);
    }

    #endregion

    #region Basic Update Operations

    [Fact]
    public void Execute_UpdatesData()
    {
        var workbook = CreateWorkbookWithChart();
        var chart = workbook.Worksheets[0].Charts[0];
        var initialSeriesCount = chart.NSeries.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "B1:B3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("data updated", result.Message);
        Assert.True(chart.NSeries.Count >= initialSeriesCount, "Chart should have data series");
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsChartIndex()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "B1:B3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#0", result.Message);
    }

    [Fact]
    public void Execute_ReturnsDataRange()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "C1:C5" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("C1:C5", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutDataRange_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("dataRange", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidChartIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 99 },
            { "dataRange", "A1:A3" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "dataRange", "A1:A3" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithChart()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "Cat";
        sheet.Cells["A2"].Value = "Dog";
        sheet.Cells["A3"].Value = "Bird";
        sheet.Cells["B1"].Value = 10;
        sheet.Cells["B2"].Value = 20;
        sheet.Cells["B3"].Value = 30;
        sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        sheet.Charts[0].NSeries.Add("B1:B3", true);
        return workbook;
    }

    private static Workbook CreateWorkbookWithMultipleCharts()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["A2"].Value = 20;
        sheet.Cells["B1"].Value = 30;
        sheet.Cells["B2"].Value = 40;

        sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        sheet.Charts[0].NSeries.Add("A1:A2", true);

        sheet.Charts.Add(ChartType.Column, 25, 0, 40, 8);
        sheet.Charts[1].NSeries.Add("B1:B2", true);

        return workbook;
    }

    #endregion
}
