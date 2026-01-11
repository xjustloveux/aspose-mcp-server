using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Chart;

public class AddExcelChartHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelChartHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_AddsToCorrectSheet()
    {
        var workbook = CreateWorkbookWithChartData();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = 1;
        workbook.Worksheets[1].Cells["A2"].Value = 2;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:A2" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(workbook.Worksheets[0].Charts);
        Assert.Single(workbook.Worksheets[1].Charts);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithChartData()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["A2"].Value = "A";
        sheet.Cells["B2"].Value = 10;
        sheet.Cells["A3"].Value = "B";
        sheet.Cells["B3"].Value = 20;
        return workbook;
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsChart()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:B3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Chart added", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDataRange()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:B3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1:B3", result);
    }

    [Fact]
    public void Execute_CreatesChartOnWorksheet()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:B3" }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].Charts);
    }

    #endregion

    #region Chart With Options

    [Fact]
    public void Execute_WithTitle_SetsTitle()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:B3" },
            { "title", "Sales Chart" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Sales Chart", workbook.Worksheets[0].Charts[0].Title.Text);
    }

    [Theory]
    [InlineData("Bar", ChartType.Bar)]
    [InlineData("Line", ChartType.Line)]
    [InlineData("Pie", ChartType.Pie)]
    [InlineData("Column", ChartType.Column)]
    public void Execute_WithChartType_SetsCorrectChartType(string chartTypeStr, ChartType expectedType)
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:B3" },
            { "chartType", chartTypeStr }
        });

        _handler.Execute(context, parameters);

        var chart = workbook.Worksheets[0].Charts[0];
        Assert.Equal(expectedType, chart.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCategoryAxisDataRange_IncludesInResult()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "B1:B3" },
            { "categoryAxisDataRange", "A1:A3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("X-axis: A1:A3", result);
    }

    [Fact]
    public void Execute_WithPosition_CreatesAtSpecifiedLocation()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:B3" },
            { "topRow", 10 },
            { "leftColumn", 5 }
        });

        _handler.Execute(context, parameters);

        var chart = workbook.Worksheets[0].Charts[0];
        Assert.Equal(10, chart.ChartObject.UpperLeftRow);
        Assert.Equal(5, chart.ChartObject.UpperLeftColumn);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutDataRange_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("dataRange", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithChartData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:B3" },
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
