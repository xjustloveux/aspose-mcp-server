using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Chart;

public class DeleteExcelChartHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelChartHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Default Values

    [Fact]
    public void Execute_DefaultsToFirstChart()
    {
        var workbook = CreateWorkbookWithCharts(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#0", result.Message);
    }

    #endregion

    #region Helper Methods

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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesChart()
    {
        var workbook = CreateWorkbookWithCharts(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsChartIndex()
    {
        var workbook = CreateWorkbookWithCharts(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#1", result.Message);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var workbook = CreateWorkbookWithCharts(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("2 remaining", result.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesAtVariousIndices(int chartIndex)
    {
        var workbook = CreateWorkbookWithCharts(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", chartIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
    }

    [Fact]
    public void Execute_ReducesChartCount()
    {
        var workbook = CreateWorkbookWithCharts(3);
        var initialCount = workbook.Worksheets[0].Charts.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, workbook.Worksheets[0].Charts.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidChartIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithCharts(1);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_NoCharts_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithCharts(1);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "chartIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
