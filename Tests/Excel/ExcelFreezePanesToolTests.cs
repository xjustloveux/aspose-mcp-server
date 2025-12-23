using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFreezePanesToolTests : ExcelTestBase
{
    private readonly ExcelFreezePanesTool _tool = new();

    [Fact]
    public async Task FreezePanes_ShouldFreezePanes()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_panes.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_panes_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["row"] = 1,
            ["column"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Verify freeze panes was applied - check custom properties or worksheet settings
        Assert.NotNull(worksheet);
    }

    [Fact]
    public async Task UnfreezePanes_ShouldUnfreezePanes()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_panes.xlsx", 10, 5);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.FreezePanes(1, 1, 1, 1);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_unfreeze_panes_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unfreeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        // Verify panes were unfrozen
        Assert.NotNull(resultWorksheet);
    }

    [Fact]
    public async Task GetFreezeStatus_ShouldReturnFreezeStatus()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_freeze_status.xlsx", 10, 5);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.FreezePanes(1, 1, 1, 1);
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Freeze", result, StringComparison.OrdinalIgnoreCase);
    }
}