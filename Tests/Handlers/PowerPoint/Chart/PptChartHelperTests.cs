using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Handlers.PowerPoint.Chart;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Chart;

public class PptChartHelperTests
{
    #region ParseChartType Tests

    [Fact]
    public void ParseChartType_WithNull_ReturnsDefault()
    {
        var result = PptChartHelper.ParseChartType(null);

        Assert.Equal(ChartType.ClusteredColumn, result);
    }

    [Fact]
    public void ParseChartType_WithEmpty_ReturnsDefault()
    {
        var result = PptChartHelper.ParseChartType("");

        Assert.Equal(ChartType.ClusteredColumn, result);
    }

    [Theory]
    [InlineData("column", ChartType.ClusteredColumn)]
    [InlineData("COLUMN", ChartType.ClusteredColumn)]
    [InlineData("bar", ChartType.ClusteredBar)]
    [InlineData("BAR", ChartType.ClusteredBar)]
    [InlineData("line", ChartType.Line)]
    [InlineData("pie", ChartType.Pie)]
    [InlineData("area", ChartType.Area)]
    [InlineData("scatter", ChartType.ScatterWithSmoothLines)]
    [InlineData("doughnut", ChartType.Doughnut)]
    [InlineData("bubble", ChartType.Bubble)]
    [InlineData("radar", ChartType.Radar)]
    [InlineData("treemap", ChartType.Treemap)]
    public void ParseChartType_WithValidValues_ReturnsCorrectType(string input, ChartType expected)
    {
        var result = PptChartHelper.ParseChartType(input);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ParseChartType_WithInvalidValue_ReturnsDefault()
    {
        var result = PptChartHelper.ParseChartType("invalid");

        Assert.Equal(ChartType.ClusteredColumn, result);
    }

    [Fact]
    public void ParseChartType_WithCustomDefault_ReturnsCustomDefault()
    {
        var result = PptChartHelper.ParseChartType("invalid", ChartType.Pie);

        Assert.Equal(ChartType.Pie, result);
    }

    #endregion

    #region GetChartByIndex Tests

    [Fact]
    public void GetChartByIndex_WithNoCharts_ThrowsArgumentException()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        var ex = Assert.Throws<ArgumentException>(() =>
            PptChartHelper.GetChartByIndex(slide, 0, 0));

        Assert.Contains("contains no charts", ex.Message);
    }

    [Fact]
    public void GetChartByIndex_WithValidIndex_ReturnsChart()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);

        var result = PptChartHelper.GetChartByIndex(slide, 0, 0);

        Assert.NotNull(result);
    }

    [Fact]
    public void GetChartByIndex_WithNegativeIndex_ThrowsArgumentException()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);

        var ex = Assert.Throws<ArgumentException>(() =>
            PptChartHelper.GetChartByIndex(slide, -1, 0));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetChartByIndex_WithIndexTooLarge_ThrowsArgumentException()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);

        var ex = Assert.Throws<ArgumentException>(() =>
            PptChartHelper.GetChartByIndex(slide, 5, 0));

        Assert.Contains("out of range", ex.Message);
        Assert.Contains("Total charts: 1", ex.Message);
    }

    #endregion

    #region SetChartTitle Tests

    [Fact]
    public void SetChartTitle_SetsTitle()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);

        PptChartHelper.SetChartTitle(chart, "Test Title");

        Assert.True(chart.HasTitle);
    }

    [Fact]
    public void SetChartTitle_WithExistingTitle_UpdatesTitle()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
        PptChartHelper.SetChartTitle(chart, "Original Title");

        PptChartHelper.SetChartTitle(chart, "Updated Title");

        Assert.True(chart.HasTitle);
    }

    #endregion
}
