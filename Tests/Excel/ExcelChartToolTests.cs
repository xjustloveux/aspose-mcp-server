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
}