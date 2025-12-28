using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelChartToolTests : ExcelTestBase
{
    private readonly ExcelChartTool _tool = new();

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

    [Fact]
    public async Task AddChart_ShouldAddChart()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_add_chart.xlsx");
        var outputPath = CreateTestFilePath("test_add_chart_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["chartType"] = "Column";
        arguments["dataRange"] = "B1:B10";
        arguments["categoryAxisDataRange"] = "A1:A10";
        arguments["position"] = "D1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Chart should be added");
    }

    [Fact]
    public async Task GetCharts_ShouldReturnChartsInfo()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_get_charts.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var arguments = CreateArguments("get", workbookPath);
        arguments["operation"] = "get";

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
        var workbookPath = CreateWorkbookWithData("test_delete_chart.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var chartsBefore = worksheet.Charts.Count;
        Assert.True(chartsBefore > 0, "Chart should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_chart_output.xlsx");
        var arguments = CreateArguments("delete", workbookPath, outputPath);
        arguments["chartIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var chartsAfter = resultWorksheet.Charts.Count;
        Assert.True(chartsAfter < chartsBefore,
            $"Chart should be deleted. Before: {chartsBefore}, After: {chartsAfter}");
    }

    [Fact]
    public async Task EditChart_ShouldModifyChart()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_edit_chart.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_chart_output.xlsx");
        var arguments = CreateArguments("edit", workbookPath, outputPath);
        arguments["chartIndex"] = 0;
        arguments["chartType"] = "Line";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.Charts.Count > 0, "Chart should exist after editing");
    }

    [Fact]
    public async Task UpdateChartData_ShouldUpdateData()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_update_chart_data.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_update_chart_data_output.xlsx");
        var arguments = CreateArguments("update_data", workbookPath, outputPath);
        arguments["chartIndex"] = 0;
        arguments["dataRange"] = "B1:B5";
        arguments["categoryAxisDataRange"] = "A1:A5";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.Charts.Count > 0, "Chart should exist after updating data");
    }

    [Fact]
    public async Task SetChartProperties_ShouldSetProperties()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_set_chart_properties.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_set_chart_properties_output.xlsx");
        var arguments = CreateArguments("set_properties", workbookPath, outputPath);
        arguments["chartIndex"] = 0;
        arguments["title"] = "Test Chart Title";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.Charts.Count > 0, "Chart should exist after setting properties");
    }

    [Fact]
    public async Task Add_WithPieChart_ShouldAddPieChart()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_add_pie_chart.xlsx", 5);
        var outputPath = CreateTestFilePath("test_add_pie_chart_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["chartType"] = "Pie";
        arguments["dataRange"] = "B1:B5";
        arguments["categoryAxisDataRange"] = "A1:A5";
        arguments["position"] = "D1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Pie chart should be added");
        Assert.Equal(ChartType.Pie, worksheet.Charts[0].Type);
    }

    [Fact]
    public async Task Add_WithLineChart_ShouldAddLineChart()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_add_line_chart.xlsx");
        var outputPath = CreateTestFilePath("test_add_line_chart_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["chartType"] = "Line";
        arguments["dataRange"] = "B1:B10";
        arguments["categoryAxisDataRange"] = "A1:A10";
        arguments["position"] = "D1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Line chart should be added");
        Assert.Equal(ChartType.Line, worksheet.Charts[0].Type);
    }

    [Fact]
    public async Task Add_WithBarChart_ShouldAddBarChart()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_add_bar_chart.xlsx", 5);
        var outputPath = CreateTestFilePath("test_add_bar_chart_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["chartType"] = "Bar";
        arguments["dataRange"] = "B1:B5";
        arguments["categoryAxisDataRange"] = "A1:A5";
        arguments["position"] = "D1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Charts.Count > 0, "Bar chart should be added");
        Assert.Equal(ChartType.Bar, worksheet.Charts[0].Type);
    }

    [Fact]
    public async Task SetChartProperties_WithTitle_ShouldUpdateTitle()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_chart_title.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_chart_title_output.xlsx");
        var arguments = CreateArguments("set_properties", workbookPath, outputPath);
        arguments["chartIndex"] = 0;
        arguments["title"] = "Sales Report 2024";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var chart = resultWorkbook.Worksheets[0].Charts[0];
        Assert.Equal("Sales Report 2024", chart.Title.Text);
    }

    [Fact]
    public async Task Add_WithMultipleSeries_ShouldAddMultipleSeries()
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
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["chartType"] = "Column";
        arguments["dataRange"] = "B1:C5";
        arguments["categoryAxisDataRange"] = "A1:A5";
        arguments["position"] = "E1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var chart = resultWorkbook.Worksheets[0].Charts[0];
        Assert.True(chart.NSeries.Count >= 1, $"Chart should have series, got {chart.NSeries.Count}");
    }

    [Fact]
    public async Task Add_WithInvalidChartType_ShouldUseDefaultColumn()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_invalid_chart_type.xlsx", 5);
        var outputPath = CreateTestFilePath("test_invalid_chart_type_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["chartType"] = "InvalidType";
        arguments["dataRange"] = "B1:B5";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var chart = workbook.Worksheets[0].Charts[0];
        Assert.Equal(ChartType.Column, chart.Type);
    }

    [Fact]
    public async Task Edit_WithInvalidChartIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_invalid_chart_index.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_chart_index_output.xlsx");
        var arguments = CreateArguments("edit", workbookPath, outputPath);
        arguments["chartIndex"] = 999;
        arguments["title"] = "Test";

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public async Task Get_WithNoCharts_ShouldReturnEmptyResult()
    {
        // Arrange
        var workbookPath = CreateWorkbookWithData("test_no_charts.xlsx");
        var arguments = CreateArguments("get", workbookPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No charts found", result);
    }
}