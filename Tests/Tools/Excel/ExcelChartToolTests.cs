using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelChartTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelChartToolTests : ExcelTestBase
{
    private readonly ExcelChartTool _tool;

    public ExcelChartToolTests()
    {
        _tool = new ExcelChartTool(SessionManager);
    }

    private string CreateWorkbookWithChartData(string fileName, int rowCount = 10, int columnCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        for (var i = 0; i < rowCount; i++)
        {
            worksheet.Cells[i, 0].Value = $"Category{i + 1}";
            for (var j = 1; j <= columnCount; j++)
                worksheet.Cells[i, j].Value = (i + 1) * 10 * j;
        }

        workbook.Save(filePath);
        return filePath;
    }

    private string CreateWorkbookWithChart(string fileName, ChartType chartType = ChartType.Column)
    {
        var filePath = CreateWorkbookWithChartData(fileName);
        using var workbook = new Workbook(filePath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(chartType, 12, 0, 27, 10);
        workbook.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddChart()
    {
        var workbookPath = CreateWorkbookWithChartData("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, dataRange: "B1:B10", categoryAxisDataRange: "A1:A10",
            outputPath: outputPath);
        Assert.StartsWith("Chart added", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Charts);
    }

    [Fact]
    public void Get_ShouldReturnChartsInfo()
    {
        var workbookPath = CreateWorkbookWithChart("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.True(json.RootElement.GetProperty("items").GetArrayLength() > 0);
    }

    [Fact]
    public void Edit_ShouldModifyChartType()
    {
        var workbookPath = CreateWorkbookWithChart("test_edit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, chartIndex: 0, chartType: "Bar", outputPath: outputPath);
        Assert.StartsWith("Chart #0 edited", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(ChartType.Bar, workbook.Worksheets[0].Charts[0].Type);
    }

    [Fact]
    public void Delete_ShouldDeleteChart()
    {
        var workbookPath = CreateWorkbookWithChart("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, chartIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Chart #0", result);
        Assert.Contains("deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Charts);
    }

    [Fact]
    public void UpdateData_ShouldUpdateChartDataRange()
    {
        var workbookPath = CreateWorkbookWithChart("test_update_data.xlsx");
        var outputPath = CreateTestFilePath("test_update_data_output.xlsx");
        var result = _tool.Execute("update_data", workbookPath, chartIndex: 0, dataRange: "B1:B5",
            categoryAxisDataRange: "A1:A5", outputPath: outputPath);
        Assert.StartsWith("Chart #0 data updated", result);
        Assert.Contains("B1:B5", result);
    }

    [Fact]
    public void SetProperties_WithTitle_ShouldSetTitle()
    {
        var workbookPath = CreateWorkbookWithChart("test_set_props_title.xlsx");
        var outputPath = CreateTestFilePath("test_set_props_title_output.xlsx");
        var result = _tool.Execute("set_properties", workbookPath, chartIndex: 0, title: "New Title",
            outputPath: outputPath);
        Assert.StartsWith("Chart #0 properties updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("New Title", workbook.Worksheets[0].Charts[0].Title.Text);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateWorkbookWithChartData($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, dataRange: "B1:B10", outputPath: outputPath);
        Assert.StartsWith("Chart added", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithChartData("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateWorkbookWithChartData("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, dataRange: "B1:B10", categoryAxisDataRange: "A1:A10");
        Assert.StartsWith("Chart added", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].Charts);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithChart("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithChart("test_session_edit.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, chartIndex: 0, chartType: "Bar");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(ChartType.Bar, workbook.Worksheets[0].Charts[0].Type);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithChart("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("delete", sessionId: sessionId, chartIndex: 0);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Charts);
    }

    [Fact]
    public void UpdateData_WithSessionId_ShouldUpdateInMemory()
    {
        var workbookPath = CreateWorkbookWithChart("test_session_update.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("update_data", sessionId: sessionId, chartIndex: 0, dataRange: "B1:B5",
            categoryAxisDataRange: "A1:A5");
        Assert.StartsWith("Chart #0 data updated", result);
    }

    [Fact]
    public void SetProperties_WithSessionId_ShouldSetPropertiesInMemory()
    {
        var workbookPath = CreateWorkbookWithChart("test_session_setprops.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("set_properties", sessionId: sessionId, chartIndex: 0, title: "Session Chart Title");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Session Chart Title", workbook.Worksheets[0].Charts[0].Title.Text);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateWorkbookWithChartData("test_path_file.xlsx");
        var sessionWorkbook = CreateWorkbookWithChart("test_session_file.xlsx");
        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion
}
