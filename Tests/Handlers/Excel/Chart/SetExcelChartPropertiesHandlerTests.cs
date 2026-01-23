using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Chart;

public class SetExcelChartPropertiesHandlerTests : ExcelHandlerTestBase
{
    private readonly SetExcelChartPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetProperties()
    {
        Assert.Equal("set_properties", _handler.Operation);
    }

    #endregion

    #region No Changes

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChangesMessage()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("no changes", result.Message);
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

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsProperties()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("properties updated", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsChartIndex()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Updated" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#0", result.Message);
    }

    #endregion

    #region Set Title

    [Fact]
    public void Execute_WithTitle_SetsTitle()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Sales Report" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Sales Report", workbook.Worksheets[0].Charts[0].Title.Text);
    }

    [Fact]
    public void Execute_WithTitle_ReturnsTitle()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "My Chart" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Title: My Chart", result.Message);
    }

    [Fact]
    public void Execute_WithRemoveTitle_RemovesTitle()
    {
        var workbook = CreateWorkbookWithChart();
        workbook.Worksheets[0].Charts[0].Title.Text = "Original Title";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "removeTitle", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Title removed", result.Message);
        Assert.Equal("", workbook.Worksheets[0].Charts[0].Title.Text);
    }

    #endregion

    #region Set Legend

    [Fact]
    public void Execute_WithLegendVisibleFalse_HidesLegend()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "legendVisible", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Legend: hide", result.Message);
    }

    [Fact]
    public void Execute_WithLegendVisibleTrue_ShowsLegend()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "legendVisible", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Legend: show", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidChartIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 99 },
            { "title", "New Title" }
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
            { "title", "New Title" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
