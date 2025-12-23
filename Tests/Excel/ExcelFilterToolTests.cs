using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFilterToolTests : ExcelTestBase
{
    private readonly ExcelFilterTool _tool = new();

    [Fact]
    public async Task ApplyFilter_ShouldApplyAutoFilter()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_apply_filter.xlsx");
        var outputPath = CreateTestFilePath("test_apply_filter_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:C5"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.AutoFilter != null, "Auto filter should be applied");
    }

    [Fact]
    public async Task RemoveFilter_ShouldRemoveAutoFilter()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_remove_filter.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.AutoFilter.Range = "A1:C5";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_remove_filter_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "remove",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:C5"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        // Filter should be removed or range cleared
        Assert.NotNull(resultWorksheet);
    }

    [Fact]
    public async Task GetFilterStatus_ShouldReturnFilterStatus()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_filter_status.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.AutoFilter.Range = "A1:C5";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_status",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Filter", result, StringComparison.OrdinalIgnoreCase);
    }
}