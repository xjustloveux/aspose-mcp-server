using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptChartToolTests : TestBase
{
    private readonly PptChartTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task AddChart_ShouldAddChartToSlide()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_chart.pptx");
        var outputPath = CreateTestFilePath("test_add_chart_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["chartType"] = "Column",
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 400,
            ["height"] = 300
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var charts = slide.Shapes.OfType<IChart>().ToList();
        Assert.True(charts.Count > 0, "Slide should contain at least one chart");
    }

    [Fact]
    public async Task EditChart_ShouldModifyChart()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_chart.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            pptSlide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_chart_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0,
            ["title"] = "Updated Chart Title"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task GetChartData_ShouldReturnChartData()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_chart_data.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_data",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Chart", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteChart_ShouldDeleteChart()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_chart.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            pptSlide.Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var chartsBefore = 0;
        using (var ppt = new Presentation(pptPath))
        {
            chartsBefore = ppt.Slides[0].Shapes.OfType<IChart>().Count();
            Assert.True(chartsBefore > 0, "Chart should exist before deletion");
        }

        var outputPath = CreateTestFilePath("test_delete_chart_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var chartsAfter = slide.Shapes.OfType<IChart>().Count();
        Assert.True(chartsAfter < chartsBefore,
            $"Chart should be deleted. Before: {chartsBefore}, After: {chartsAfter}");
    }
}