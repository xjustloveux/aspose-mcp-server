using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelChartToolTests : ExcelTestBase
{
    private readonly ExcelChartTool _tool;

    public ExcelChartToolTests()
    {
        _tool = new ExcelChartTool(SessionManager);
    }

    private string CreateWorkbookWithData(string fileName, int rowCount = 10)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Add sample data
        for (var i = 0; i < rowCount; i++)
        {
            worksheet.Cells[i, 0].Value = $"Category{i + 1}";
            worksheet.Cells[i, 1].Value = (i + 1) * 10;
        }

        workbook.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddChart_ShouldAddChart()
    {
        var workbookPath = CreateWorkbookWithData("test_add_chart.xlsx");
        var outputPath = CreateTestFilePath("test_add_chart_output.xlsx");
        _tool.Execute("add", workbookPath, chartType: "Column", dataRange: "B1:B10", categoryAxisDataRange: "A1:A10",
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Chart should be added");
    }

    [Fact]
    public void GetCharts_ShouldReturnChartsInfo()
    {
        var workbookPath = CreateWorkbookWithData("test_get_charts.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);
        var result = _tool.Execute("get", workbookPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Chart", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteChart_ShouldDeleteChart()
    {
        var workbookPath = CreateWorkbookWithData("test_delete_chart.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var chartsBefore = worksheet.Charts.Count;
        Assert.True(chartsBefore > 0, "Chart should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_chart_output.xlsx");
        _tool.Execute("delete", workbookPath, chartIndex: 0, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var chartsAfter = resultWorksheet.Charts.Count;
        Assert.True(chartsAfter < chartsBefore,
            $"Chart should be deleted. Before: {chartsBefore}, After: {chartsAfter}");
    }

    [Fact]
    public void EditChart_ShouldModifyChart()
    {
        var workbookPath = CreateWorkbookWithData("test_edit_chart.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_chart_output.xlsx");
        _tool.Execute("edit", workbookPath, chartIndex: 0, chartType: "Line", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.Charts.Count > 0, "Chart should exist after editing");
    }

    [Fact]
    public void UpdateChartData_ShouldUpdateData()
    {
        var workbookPath = CreateWorkbookWithData("test_update_chart_data.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_update_chart_data_output.xlsx");
        _tool.Execute("update_data", workbookPath, chartIndex: 0, dataRange: "B1:B5", categoryAxisDataRange: "A1:A5",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.Charts.Count > 0, "Chart should exist after updating data");
    }

    [Fact]
    public void SetChartProperties_ShouldSetProperties()
    {
        var workbookPath = CreateWorkbookWithData("test_set_chart_properties.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_set_chart_properties_output.xlsx");
        _tool.Execute("set_properties", workbookPath, chartIndex: 0, title: "Test Chart Title", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.Charts.Count > 0, "Chart should exist after setting properties");
    }

    [Fact]
    public void Add_WithPieChart_ShouldAddPieChart()
    {
        var workbookPath = CreateWorkbookWithData("test_add_pie_chart.xlsx", 5);
        var outputPath = CreateTestFilePath("test_add_pie_chart_output.xlsx");
        _tool.Execute("add", workbookPath, chartType: "Pie", dataRange: "B1:B5", categoryAxisDataRange: "A1:A5",
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Pie chart should be added");
        Assert.Equal(ChartType.Pie, worksheet.Charts[0].Type);
    }

    [Fact]
    public void Add_WithLineChart_ShouldAddLineChart()
    {
        var workbookPath = CreateWorkbookWithData("test_add_line_chart.xlsx");
        var outputPath = CreateTestFilePath("test_add_line_chart_output.xlsx");
        _tool.Execute("add", workbookPath, chartType: "Line", dataRange: "B1:B10", categoryAxisDataRange: "A1:A10",
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Line chart should be added");
        Assert.Equal(ChartType.Line, worksheet.Charts[0].Type);
    }

    [Fact]
    public void Add_WithBarChart_ShouldAddBarChart()
    {
        var workbookPath = CreateWorkbookWithData("test_add_bar_chart.xlsx", 5);
        var outputPath = CreateTestFilePath("test_add_bar_chart_output.xlsx");
        _tool.Execute("add", workbookPath, chartType: "Bar", dataRange: "B1:B5", categoryAxisDataRange: "A1:A5",
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Bar chart should be added");
        Assert.Equal(ChartType.Bar, worksheet.Charts[0].Type);
    }

    [Fact]
    public void SetChartProperties_WithTitle_ShouldUpdateTitle()
    {
        var workbookPath = CreateWorkbookWithData("test_chart_title.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_chart_title_output.xlsx");
        _tool.Execute("set_properties", workbookPath, chartIndex: 0, title: "Sales Report 2024",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var chart = resultWorkbook.Worksheets[0].Charts[0];
        Assert.Equal("Sales Report 2024", chart.Title.Text);
    }

    [Fact]
    public void Add_WithMultipleSeries_ShouldAddMultipleSeries()
    {
        // Arrange - Create workbook with multiple data columns
        var workbookPath = CreateTestFilePath("test_multi_series_chart.xlsx");
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Add category labels
        for (var i = 0; i < 5; i++)
            worksheet.Cells[i, 0].Value = $"Month{i + 1}";

        // Add first series data
        for (var i = 0; i < 5; i++)
            worksheet.Cells[i, 1].Value = (i + 1) * 10;

        // Add second series data
        for (var i = 0; i < 5; i++)
            worksheet.Cells[i, 2].Value = (i + 1) * 15;

        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_multi_series_chart_output.xlsx");
        _tool.Execute("add", workbookPath, chartType: "Column", dataRange: "B1:C5", categoryAxisDataRange: "A1:A5",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var chart = resultWorkbook.Worksheets[0].Charts[0];
        Assert.True(chart.NSeries.Count >= 1, $"Chart should have series, got {chart.NSeries.Count}");
    }

    [Fact]
    public void Add_WithInvalidChartType_ShouldUseDefaultColumn()
    {
        var workbookPath = CreateWorkbookWithData("test_invalid_chart_type.xlsx", 5);
        var outputPath = CreateTestFilePath("test_invalid_chart_type_output.xlsx");
        _tool.Execute("add", workbookPath, chartType: "InvalidType", dataRange: "B1:B5", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var chart = workbook.Worksheets[0].Charts[0];
        Assert.Equal(ChartType.Column, chart.Type);
    }

    [Fact]
    public void Edit_WithInvalidChartIndex_ShouldThrowException()
    {
        var workbookPath = CreateWorkbookWithData("test_invalid_chart_index.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_chart_index_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, chartIndex: 999, title: "Test", outputPath: outputPath));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithNoCharts_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateWorkbookWithData("test_no_charts.xlsx");
        var result = _tool.Execute("get", workbookPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No charts found", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithData("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", workbookPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingDataRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithData("test_missing_data_range.xlsx");
        var outputPath = CreateTestFilePath("test_missing_data_range_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, chartType: "Column", outputPath: outputPath));

        Assert.Contains("dataRange is required", ex.Message);
    }

    [Fact]
    public void Delete_WithDefaultChartIndex_ShouldDeleteFirstChart()
    {
        var workbookPath = CreateWorkbookWithData("test_delete_default_index.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_default_index_output.xlsx");

        // Act - chartIndex defaults to 0, which is valid when chart exists
        var result = _tool.Execute("delete", workbookPath, outputPath: outputPath);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].Charts);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithData("test_session_get_charts.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("Chart", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateWorkbookWithData("test_session_add_chart.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, chartType: "Column", dataRange: "B1:B10",
            categoryAxisDataRange: "A1:A10");
        Assert.Contains("Chart added", result);

        // Verify in-memory workbook has the chart
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(sessionWorkbook.Worksheets[0].Charts.Count > 0, "Chart should be added in memory");
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithData("test_session_edit_chart.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, chartIndex: 0, chartType: "Line");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(ChartType.Line, sessionWorkbook.Worksheets[0].Charts[0].Type);
    }

    [Fact]
    public void SetProperties_WithSessionId_ShouldSetPropertiesInMemory()
    {
        var workbookPath = CreateWorkbookWithData("test_session_set_props.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("set_properties", sessionId: sessionId, chartIndex: 0, title: "Session Chart Title");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Session Chart Title", sessionWorkbook.Worksheets[0].Charts[0].Title.Text);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}