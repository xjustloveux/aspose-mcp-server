using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Chart;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptChartTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptChartToolTests : PptTestBase
{
    private readonly PptChartTool _tool;

    public PptChartToolTests()
    {
        _tool = new PptChartTool(SessionManager);
    }

    private string CreatePresentationWithChart(string fileName, ChartType chartType = ChartType.ClusteredColumn)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        ppt.Slides[0].Shapes.AddChart(chartType, 100, 100, 400, 300);
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddChartToSlide()
    {
        var pptPath = CreatePresentation("test_add_chart.pptx");
        var outputPath = CreateTestFilePath("test_add_chart_output.pptx");
        var result = _tool.Execute("add", 0, pptPath, chartType: "Column", x: 100, y: 100, width: 400, height: 300,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Chart", data.Message);
        Assert.Contains("added to slide", data.Message);
        using var presentation = new Presentation(outputPath);
        var charts = presentation.Slides[0].Shapes.OfType<IChart>().ToList();
        Assert.Single(charts);
    }

    [Fact]
    public void Delete_ShouldRemoveChart()
    {
        var pptPath = CreatePresentationWithChart("test_delete_chart.pptx");
        var outputPath = CreateTestFilePath("test_delete_chart_output.pptx");
        var result = _tool.Execute("delete", 0, pptPath, shapeIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Chart", data.Message);
        Assert.Contains("deleted from slide", data.Message);
        using var presentation = new Presentation(outputPath);
        var charts = presentation.Slides[0].Shapes.OfType<IChart>().ToList();
        Assert.Empty(charts);
    }

    [Fact]
    public void GetData_ShouldReturnChartData()
    {
        var pptPath = CreatePresentationWithChart("test_get_data.pptx");
        var result = _tool.Execute("get_data", 0, pptPath, shapeIndex: 0);
        var data = GetResultData<GetChartDataPptResult>(result);
        Assert.NotNull(data.ChartType);
        Assert.NotNull(data.Categories);
        Assert.NotNull(data.Series);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, chartType: "Column", x: 100, y: 100, width: 400, height: 300,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Chart", data.Message);
        Assert.Contains("added to slide", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", 0, pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IChart>().Count();
        var result = _tool.Execute("add", 0, sessionId: sessionId, chartType: "Column", x: 100, y: 100, width: 400,
            height: 300);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Chart", data.Message);
        Assert.Contains("added to slide", data.Message);
        Assert.True(ppt.Slides[0].Shapes.OfType<IChart>().Count() > initialCount);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithChart("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IChart>().Count();
        var result = _tool.Execute("delete", 0, sessionId: sessionId, shapeIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Chart", data.Message);
        Assert.Contains("deleted from slide", data.Message);
        Assert.True(ppt.Slides[0].Shapes.OfType<IChart>().Count() < initialCount);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void GetData_WithSessionId_ShouldReturnData()
    {
        var pptPath = CreatePresentationWithChart("test_session_getdata.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_data", 0, sessionId: sessionId, shapeIndex: 0);
        var data = GetResultData<GetChartDataPptResult>(result);
        Assert.NotNull(data.ChartType);
        Assert.NotNull(data.Categories);
        var output = GetResultOutput<GetChartDataPptResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_data", 0, sessionId: "invalid_session", shapeIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithChart("test_path_chart.pptx");
        var pptPath2 = CreatePresentationWithChart("test_session_chart.pptx", ChartType.Pie);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get_data", 0, pptPath1, sessionId, shapeIndex: 0);
        var data = GetResultData<GetChartDataPptResult>(result);
        Assert.Contains("Pie", data.ChartType);
    }

    #endregion
}
