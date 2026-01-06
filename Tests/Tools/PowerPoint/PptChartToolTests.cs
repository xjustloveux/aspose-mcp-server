using System.Text.Json.Nodes;
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

    private string CreateTestPresentation(string fileName, int slideCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithChart(string fileName, ChartType chartType = ChartType.ClusteredColumn)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        ppt.Slides[0].Shapes.AddChart(chartType, 100, 100, 400, 300);
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithMultipleCharts(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        ppt.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 300, 200);
        ppt.Slides[0].Shapes.AddChart(ChartType.Pie, 400, 50, 300, 200);
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddChartToSlide()
    {
        var pptPath = CreateTestPresentation("test_add_chart.pptx");
        var outputPath = CreateTestFilePath("test_add_chart_output.pptx");
        var result = _tool.Execute("add", 0, pptPath, chartType: "Column", x: 100, y: 100, width: 400, height: 300,
            outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("added to slide", result);
        using var presentation = new Presentation(outputPath);
        var charts = presentation.Slides[0].Shapes.OfType<IChart>().ToList();
        Assert.Single(charts);
    }

    [Fact]
    public void Add_WithTitle_ShouldSetChartTitle()
    {
        var pptPath = CreateTestPresentation("test_add_chart_title.pptx");
        var outputPath = CreateTestFilePath("test_add_chart_title_output.pptx");
        _tool.Execute("add", 0, pptPath, chartType: "Bar", title: "Sales Report", x: 100, y: 100, width: 400,
            height: 300, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var chart = presentation.Slides[0].Shapes.OfType<IChart>().First();
        Assert.True(chart.HasTitle);
    }

    [Fact]
    public void Add_WithCustomPosition_ShouldUseProvidedValues()
    {
        var pptPath = CreateTestPresentation("test_add_chart_position.pptx");
        var outputPath = CreateTestFilePath("test_add_chart_position_output.pptx");
        _tool.Execute("add", 0, pptPath, chartType: "Line", x: 150, y: 200, width: 350, height: 250,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var chart = presentation.Slides[0].Shapes.OfType<IChart>().First();
        Assert.Equal(150, chart.X);
        Assert.Equal(200, chart.Y);
        Assert.Equal(350, chart.Width);
        Assert.Equal(250, chart.Height);
    }

    [Theory]
    [InlineData("Column", ChartType.ClusteredColumn)]
    [InlineData("Bar", ChartType.ClusteredBar)]
    [InlineData("Line", ChartType.Line)]
    [InlineData("Pie", ChartType.Pie)]
    [InlineData("Area", ChartType.Area)]
    public void Add_WithDifferentChartTypes_ShouldCreateCorrectType(string chartTypeStr, ChartType expectedType)
    {
        var pptPath = CreateTestPresentation($"test_add_{chartTypeStr}.pptx");
        var outputPath = CreateTestFilePath($"test_add_{chartTypeStr}_output.pptx");
        _tool.Execute("add", 0, pptPath, chartType: chartTypeStr, x: 100, y: 100, width: 400, height: 300,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var chart = presentation.Slides[0].Shapes.OfType<IChart>().First();
        Assert.Equal(expectedType, chart.Type);
    }

    [Fact]
    public void Edit_ShouldModifyChartTitle()
    {
        var pptPath = CreatePresentationWithChart("test_edit_chart.pptx");
        var outputPath = CreateTestFilePath("test_edit_chart_output.pptx");
        _tool.Execute("edit", 0, pptPath, shapeIndex: 0, title: "Updated Chart Title", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var chart = presentation.Slides[0].Shapes.OfType<IChart>().First();
        Assert.True(chart.HasTitle);
        Assert.NotNull(chart.ChartTitle);
    }

    [Fact]
    public void Edit_ShouldChangeChartTypeWithinSameFamily()
    {
        var pptPath = CreatePresentationWithChart("test_edit_type_family.pptx");
        var outputPath = CreateTestFilePath("test_edit_type_family_output.pptx");
        _tool.Execute("edit", 0, pptPath, shapeIndex: 0, chartType: "Column", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var chart = presentation.Slides[0].Shapes.OfType<IChart>().First();
        Assert.Equal(ChartType.ClusteredColumn, chart.Type);
    }

    [Fact]
    public void Edit_WithIncompatibleChartType_ShouldThrowInvalidOperationException()
    {
        var pptPath = CreatePresentationWithChart("test_edit_incompatible.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("edit", 0, pptPath, shapeIndex: 0, chartType: "Bar"));
        Assert.Contains("Failed to change chart type", ex.Message);
    }

    [Fact]
    public void Delete_ShouldRemoveChart()
    {
        var pptPath = CreatePresentationWithChart("test_delete_chart.pptx");
        var outputPath = CreateTestFilePath("test_delete_chart_output.pptx");
        var result = _tool.Execute("delete", 0, pptPath, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("deleted from slide", result);
        using var presentation = new Presentation(outputPath);
        var charts = presentation.Slides[0].Shapes.OfType<IChart>().ToList();
        Assert.Empty(charts);
    }

    [Fact]
    public void Delete_WithMultipleCharts_ShouldRemoveCorrectChart()
    {
        var pptPath = CreatePresentationWithMultipleCharts("test_delete_multiple.pptx");
        var outputPath = CreateTestFilePath("test_delete_multiple_output.pptx");
        _tool.Execute("delete", 0, pptPath, shapeIndex: 0, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var charts = presentation.Slides[0].Shapes.OfType<IChart>().ToList();
        Assert.Single(charts);
        Assert.Equal(ChartType.Pie, charts[0].Type);
    }

    [Fact]
    public void GetData_ShouldReturnChartData()
    {
        var pptPath = CreatePresentationWithChart("test_get_data.pptx");
        var result = _tool.Execute("get_data", 0, pptPath, shapeIndex: 0);
        Assert.Contains("chartType", result);
        Assert.Contains("categories", result);
        Assert.Contains("series", result);
    }

    [Fact]
    public void GetData_ShouldIncludeChartInfo()
    {
        var pptPath = CreatePresentationWithChart("test_get_data_info.pptx");
        var result = _tool.Execute("get_data", 0, pptPath, shapeIndex: 0);
        Assert.Contains("slideIndex", result);
        Assert.Contains("chartIndex", result);
        Assert.Contains("hasTitle", result);
    }

    [Fact]
    public void UpdateData_WithCategories_ShouldUpdateData()
    {
        var pptPath = CreatePresentationWithChart("test_update_data.pptx");
        var outputPath = CreateTestFilePath("test_update_data_output.pptx");
        var data = new JsonObject
        {
            ["categories"] = new JsonArray("Q1", "Q2", "Q3"),
            ["series"] = new JsonArray(new JsonObject
            {
                ["name"] = "Sales",
                ["values"] = new JsonArray(100.0, 200.0, 150.0)
            })
        };
        var result = _tool.Execute("update_data", 0, pptPath, shapeIndex: 0, data: data, outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void UpdateData_WithClearExisting_ShouldClearData()
    {
        var pptPath = CreatePresentationWithChart("test_update_clear.pptx");
        var outputPath = CreateTestFilePath("test_update_clear_output.pptx");
        var result = _tool.Execute("update_data", 0, pptPath, shapeIndex: 0, clearExisting: true,
            outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("updated", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, chartType: "Column", x: 100, y: 100, width: 400, height: 300,
            outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("added to slide", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pptPath = CreatePresentationWithChart($"test_case_edit_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, shapeIndex: 0, title: "Test", outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("updated", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pptPath = CreatePresentationWithChart($"test_case_delete_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("deleted from slide", result);
    }

    [Theory]
    [InlineData("GET_DATA")]
    [InlineData("Get_Data")]
    [InlineData("get_data")]
    public void Operation_ShouldBeCaseInsensitive_GetData(string operation)
    {
        var pptPath = CreatePresentationWithChart($"test_case_getdata_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, 0, pptPath, shapeIndex: 0);
        Assert.Contains("chartType", result);
    }

    [Theory]
    [InlineData("UPDATE_DATA")]
    [InlineData("Update_Data")]
    [InlineData("update_data")]
    public void Operation_ShouldBeCaseInsensitive_UpdateData(string operation)
    {
        var pptPath = CreatePresentationWithChart($"test_case_updatedata_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_updatedata_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, shapeIndex: 0, clearExisting: true, outputPath: outputPath);
        Assert.StartsWith("Chart", result);
        Assert.Contains("updated", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", 0, pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithoutChartType_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_type.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", 0, pptPath));
        Assert.Contains("chartType is required", ex.Message);
    }

    [Fact]
    public void Edit_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithChart("test_edit_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("edit", 0, pptPath, title: "Test"));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void Delete_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithChart("test_delete_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", 0, pptPath));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void GetData_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithChart("test_getdata_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_data", 0, pptPath));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void UpdateData_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithChart("test_updatedata_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("update_data", 0, pptPath));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void GetData_OnSlideWithNoCharts_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_no_charts.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_data", 0, pptPath, shapeIndex: 0));
        Assert.Contains("no charts", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidChartIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithChart("test_invalid_index.pptx");
        var ex =
            Assert.Throws<ArgumentException>(() => _tool.Execute("edit", 0, pptPath, shapeIndex: 99, title: "Test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", 99, pptPath, chartType: "Column"));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IChart>().Count();
        var result = _tool.Execute("add", 0, sessionId: sessionId, chartType: "Column", x: 100, y: 100, width: 400,
            height: 300);
        Assert.StartsWith("Chart", result);
        Assert.Contains("added to slide", result);
        Assert.Contains("session", result);
        Assert.True(ppt.Slides[0].Shapes.OfType<IChart>().Count() > initialCount);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentationWithChart("test_session_edit.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("edit", 0, sessionId: sessionId, shapeIndex: 0, title: "Session Chart Title");
        Assert.StartsWith("Chart", result);
        Assert.Contains("updated", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var chart = ppt.Slides[0].Shapes.OfType<IChart>().First();
        Assert.True(chart.HasTitle);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithChart("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IChart>().Count();
        var result = _tool.Execute("delete", 0, sessionId: sessionId, shapeIndex: 0);
        Assert.StartsWith("Chart", result);
        Assert.Contains("deleted from slide", result);
        Assert.True(ppt.Slides[0].Shapes.OfType<IChart>().Count() < initialCount);
    }

    [Fact]
    public void GetData_WithSessionId_ShouldReturnData()
    {
        var pptPath = CreatePresentationWithChart("test_session_getdata.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_data", 0, sessionId: sessionId, shapeIndex: 0);
        Assert.Contains("chartType", result);
        Assert.Contains("categories", result);
    }

    [Fact]
    public void UpdateData_WithSessionId_ShouldUpdateInMemory()
    {
        var pptPath = CreatePresentationWithChart("test_session_updatedata.pptx");
        var sessionId = OpenSession(pptPath);
        var data = new JsonObject
        {
            ["categories"] = new JsonArray("A", "B"),
            ["series"] = new JsonArray(new JsonObject
            {
                ["name"] = "Data",
                ["values"] = new JsonArray(10.0, 20.0)
            })
        };
        var result = _tool.Execute("update_data", 0, sessionId: sessionId, shapeIndex: 0, data: data);
        Assert.StartsWith("Chart", result);
        Assert.Contains("updated", result);
        Assert.Contains("session", result);
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
        Assert.Contains("Pie", result);
    }

    #endregion
}