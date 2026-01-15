using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Chart;

namespace AsposeMcpServer.Tests.Handlers.Excel.Chart;

public class ExcelChartHelperTests
{
    #region ParseChartType Tests

    [Fact]
    public void ParseChartType_WithNull_ReturnsDefault()
    {
        var result = ExcelChartHelper.ParseChartType(null);

        Assert.Equal(ChartType.Column, result);
    }

    [Fact]
    public void ParseChartType_WithEmpty_ReturnsDefault()
    {
        var result = ExcelChartHelper.ParseChartType("");

        Assert.Equal(ChartType.Column, result);
    }

    [Theory]
    [InlineData("Column", ChartType.Column)]
    [InlineData("column", ChartType.Column)]
    [InlineData("COLUMN", ChartType.Column)]
    [InlineData("Bar", ChartType.Bar)]
    [InlineData("Line", ChartType.Line)]
    [InlineData("Pie", ChartType.Pie)]
    [InlineData("Area", ChartType.Area)]
    public void ParseChartType_WithValidValues_ReturnsCorrectType(string input, ChartType expected)
    {
        var result = ExcelChartHelper.ParseChartType(input);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ParseChartType_WithInvalidValue_ReturnsDefault()
    {
        var result = ExcelChartHelper.ParseChartType("InvalidChart");

        Assert.Equal(ChartType.Column, result);
    }

    [Fact]
    public void ParseChartType_WithCustomDefault_ReturnsCustomDefault()
    {
        var result = ExcelChartHelper.ParseChartType("InvalidChart", ChartType.Pie);

        Assert.Equal(ChartType.Pie, result);
    }

    #endregion

    #region ParseLegendPosition Tests

    [Fact]
    public void ParseLegendPosition_WithNull_ReturnsDefault()
    {
        var result = ExcelChartHelper.ParseLegendPosition(null);

        Assert.Equal(LegendPositionType.Bottom, result);
    }

    [Fact]
    public void ParseLegendPosition_WithEmpty_ReturnsDefault()
    {
        var result = ExcelChartHelper.ParseLegendPosition("");

        Assert.Equal(LegendPositionType.Bottom, result);
    }

    [Theory]
    [InlineData("bottom", LegendPositionType.Bottom)]
    [InlineData("BOTTOM", LegendPositionType.Bottom)]
    [InlineData("Bottom", LegendPositionType.Bottom)]
    [InlineData("top", LegendPositionType.Top)]
    [InlineData("left", LegendPositionType.Left)]
    [InlineData("right", LegendPositionType.Right)]
    [InlineData("topright", LegendPositionType.Right)]
    public void ParseLegendPosition_WithValidValues_ReturnsCorrectPosition(string input,
        LegendPositionType expected)
    {
        var result = ExcelChartHelper.ParseLegendPosition(input);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ParseLegendPosition_WithInvalidValue_ReturnsDefault()
    {
        var result = ExcelChartHelper.ParseLegendPosition("center");

        Assert.Equal(LegendPositionType.Bottom, result);
    }

    [Fact]
    public void ParseLegendPosition_WithCustomDefault_ReturnsCustomDefault()
    {
        var result = ExcelChartHelper.ParseLegendPosition("invalid", LegendPositionType.Top);

        Assert.Equal(LegendPositionType.Top, result);
    }

    #endregion

    #region GetChart Tests

    [Fact]
    public void GetChart_WithValidIndex_ReturnsChart()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue(1);
        worksheet.Cells["A2"].PutValue(2);
        worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);

        var result = ExcelChartHelper.GetChart(worksheet, 0);

        Assert.NotNull(result);
    }

    [Fact]
    public void GetChart_WithNegativeIndex_ThrowsArgumentException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelChartHelper.GetChart(worksheet, -1));

        Assert.Contains("Chart index -1 is out of range", ex.Message);
    }

    [Fact]
    public void GetChart_WithIndexTooLarge_ThrowsArgumentException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelChartHelper.GetChart(worksheet, 5));

        Assert.Contains("Chart index 5 is out of range", ex.Message);
        Assert.Contains("worksheet has 1 charts", ex.Message);
    }

    [Fact]
    public void GetChart_WithNoCharts_ThrowsArgumentException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelChartHelper.GetChart(worksheet, 0));

        Assert.Contains("worksheet has 0 charts", ex.Message);
    }

    #endregion

    #region SetCategoryData Tests

    [Fact]
    public void SetCategoryData_WithEmptyRange_DoesNothing()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Cat1");
        worksheet.Cells["B1"].PutValue(1);
        var chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B1:B5", true);

        ExcelChartHelper.SetCategoryData(chart, "");
    }

    [Fact]
    public void SetCategoryData_WithNullRange_DoesNothing()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B1:B5", true);

        ExcelChartHelper.SetCategoryData(chart, null!);
    }

    [Fact]
    public void SetCategoryData_WithNoSeries_DoesNothing()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];

        ExcelChartHelper.SetCategoryData(chart, "A1:A5");
    }

    [Fact]
    public void SetCategoryData_WithValidRange_SetsCategoryData()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Cat1");
        worksheet.Cells["B1"].PutValue(1);
        var chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B1:B5", true);

        ExcelChartHelper.SetCategoryData(chart, "A1:A5");

        Assert.NotNull(chart.NSeries.CategoryData);
        Assert.NotEmpty(chart.NSeries.CategoryData);
    }

    #endregion

    #region AddDataSeries Tests

    [Fact]
    public void AddDataSeries_WithSingleRange_AddsSeries()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue(1);
        worksheet.Cells["A2"].PutValue(2);
        var chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];

        ExcelChartHelper.AddDataSeries(chart, "A1:A5");

        Assert.Single(chart.NSeries);
    }

    [Fact]
    public void AddDataSeries_WithMultipleRanges_AddsMultipleSeries()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue(1);
        worksheet.Cells["B1"].PutValue(2);
        var chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];

        ExcelChartHelper.AddDataSeries(chart, "A1:A5, B1:B5");

        Assert.Equal(2, chart.NSeries.Count);
    }

    [Fact]
    public void AddDataSeries_ClearsPreviousSeries()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue(1);
        var chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("C1:C5", true);

        ExcelChartHelper.AddDataSeries(chart, "A1:A5");

        Assert.Single(chart.NSeries);
    }

    #endregion
}
