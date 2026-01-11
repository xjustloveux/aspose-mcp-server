using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Chart;

public class EditExcelChartHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelChartHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region No Changes

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChangesMessage()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("no changes", result);
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

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsChart()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 },
            { "title", "New Title" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsChartIndex()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 },
            { "title", "Updated" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#0", result);
    }

    #endregion

    #region Edit Title

    [Fact]
    public void Execute_WithTitle_ChangesTitle()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 },
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
            { "chartIndex", 0 },
            { "title", "My Chart" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Title: My Chart", result);
    }

    #endregion

    #region Edit Legend

    [Fact]
    public void Execute_WithShowLegendFalse_HidesLegend()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 },
            { "showLegend", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Legend: hide", result);
    }

    [Fact]
    public void Execute_WithShowLegendTrue_ShowsLegend()
    {
        var workbook = CreateWorkbookWithChart();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 },
            { "showLegend", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Legend: show", result);
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
            { "chartIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
