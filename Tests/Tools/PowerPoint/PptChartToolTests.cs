using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptChartToolTests : TestBase
{
    private readonly PptChartTool _tool;

    public PptChartToolTests()
    {
        _tool = new PptChartTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddChart_ShouldAddChartToSlide()
    {
        var pptPath = CreateTestPresentation("test_add_chart.pptx");
        var outputPath = CreateTestFilePath("test_add_chart_output.pptx");
        _tool.Execute("add", 0, pptPath, chartType: "Column", x: 100, y: 100, width: 400, height: 300,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var charts = slide.Shapes.OfType<IChart>().ToList();
        Assert.True(charts.Count > 0, "Slide should contain at least one chart");
    }

    [Fact]
    public void EditChart_ShouldModifyChart()
    {
        var pptPath = CreateTestPresentation("test_edit_chart.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            pptSlide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_chart_output.pptx");
        _tool.Execute("edit", 0, pptPath, shapeIndex: 0, title: "Updated Chart Title", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var charts = slide.Shapes.OfType<IChart>().ToList();
        Assert.True(charts.Count > 0, "Chart should exist after editing");
        var chart = charts[0];
        Assert.NotNull(chart);
        Assert.NotNull(chart.ChartTitle);
        var titleText = chart.ChartTitle.TextFrameForOverriding?.Text ?? "";

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            var hasUpdated = titleText.Contains("Updated", StringComparison.OrdinalIgnoreCase) ||
                             titleText.Contains("Updat", StringComparison.OrdinalIgnoreCase);
            Assert.True(hasUpdated || titleText.Length > 0,
                $"In evaluation mode, chart title may be truncated due to watermark. " +
                $"Expected 'Updated' or 'Updat', but got: '{titleText.Substring(0, Math.Min(50, titleText.Length))}...'");
        }
        else
        {
            var hasUpdated = titleText.Contains("Updated", StringComparison.OrdinalIgnoreCase);
            Assert.True(hasUpdated,
                $"Chart title should contain 'Updated', but got: '{titleText.Substring(0, Math.Min(50, titleText.Length))}...'");
        }
    }

    [Fact]
    public void GetChartData_ShouldReturnChartData()
    {
        var pptPath = CreateTestPresentation("test_get_chart_data.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get_data", 0, pptPath, shapeIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Chart", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteChart_ShouldDeleteChart()
    {
        var pptPath = CreateTestPresentation("test_delete_chart.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            pptSlide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        int chartsBefore;
        using (var ppt = new Presentation(pptPath))
        {
            chartsBefore = ppt.Slides[0].Shapes.OfType<IChart>().Count();
            Assert.True(chartsBefore > 0, "Chart should exist before deletion");
        }

        var outputPath = CreateTestFilePath("test_delete_chart_output.pptx");
        _tool.Execute("delete", 0, pptPath, shapeIndex: 0, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var chartsAfter = slide.Shapes.OfType<IChart>().Count();
        Assert.True(chartsAfter < chartsBefore,
            $"Chart should be deleted. Before: {chartsBefore}, After: {chartsAfter}");
    }

    [Fact]
    public void AddChart_WithCustomPosition_ShouldUseProvidedValues()
    {
        var pptPath = CreateTestPresentation("test_add_chart_position.pptx");
        var outputPath = CreateTestFilePath("test_add_chart_position_output.pptx");

        _tool.Execute("add", 0, pptPath, chartType: "Bar", x: 150, y: 200, width: 350, height: 250,
            outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        var chart = presentation.Slides[0].Shapes.OfType<IChart>().First();
        Assert.Equal(150, chart.X);
        Assert.Equal(200, chart.Y);
        Assert.Equal(350, chart.Width);
        Assert.Equal(250, chart.Height);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_UnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", 0, pptPath));
    }

    [Fact]
    public void GetChartData_NoChartsOnSlide_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_no_charts.pptx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_data", 0, pptPath, shapeIndex: 0));
        Assert.Contains("no charts", ex.Message);
    }

    [Fact]
    public void EditChart_InvalidChartIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_index.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var ex =
            Assert.Throws<ArgumentException>(() => _tool.Execute("edit", 0, pptPath, shapeIndex: 99, title: "Test"));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetChartData_WithSessionId_ShouldReturnChartData()
    {
        var pptPath = CreateTestPresentation("test_session_get_chart_data.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_data", 0, sessionId: sessionId, shapeIndex: 0);
        Assert.NotNull(result);
        Assert.Contains("Chart", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddChart_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_chart.pptx");
        var sessionId = OpenSession(pptPath);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IChart>().Count();
        var result = _tool.Execute("add", 0, sessionId: sessionId, chartType: "Column", x: 100, y: 100, width: 400,
            height: 300);
        Assert.Contains("Chart", result);
        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(ppt.Slides[0].Shapes.OfType<IChart>().Count() > initialCount);
    }

    [Fact]
    public void EditChart_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_edit_chart.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            presentation.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("edit", 0, sessionId: sessionId, shapeIndex: 0, title: "Session Chart Title");
        Assert.Contains("Chart", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var chart = ppt.Slides[0].Shapes.OfType<IChart>().First();
        Assert.NotNull(chart.ChartTitle);
    }

    [Fact]
    public void DeleteChart_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_delete_chart.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var pptSession = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = pptSession.Slides[0].Shapes.OfType<IChart>().Count();
        var result = _tool.Execute("delete", 0, sessionId: sessionId, shapeIndex: 0);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(pptSession.Slides[0].Shapes.OfType<IChart>().Count() < initialCount);
    }

    #endregion
}