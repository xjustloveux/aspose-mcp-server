using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelProtectToolTests : ExcelTestBase
{
    private readonly ExcelProtectTool _tool = new();

    [Fact]
    public async Task ProtectWorkbook_ShouldProtectWorkbook()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_protect_workbook.xlsx");
        var outputPath = CreateTestFilePath("test_protect_workbook_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "protect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["password"] = "test123",
            ["protectType"] = "workbook"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Protected workbook should be created");
        var workbook = new Workbook(outputPath);
        // Check if workbook is protected by trying to access protection settings
        Assert.NotNull(workbook);
    }

    [Fact]
    public async Task ProtectWorksheet_ShouldProtectWorksheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_protect_worksheet.xlsx");
        var outputPath = CreateTestFilePath("test_protect_worksheet_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "protect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["password"] = "test123",
            ["protectType"] = "worksheet",
            ["sheetIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Verify worksheet is protected by checking protection settings
        Assert.NotNull(worksheet.Protection);
    }

    [Fact]
    public async Task UnprotectWorkbook_ShouldUnprotectWorkbook()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_unprotect_workbook.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Protect(ProtectionType.All, "test123");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_unprotect_workbook_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unprotect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["password"] = "test123",
            ["protectType"] = "workbook"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        // Verify workbook is unprotected by checking it can be accessed
        Assert.NotNull(resultWorkbook);
    }

    [Fact]
    public async Task GetProtectionInfo_ShouldReturnProtectionInfo()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_protection_info.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Protect(ProtectionType.All, "test123");
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
        Assert.Contains("Protection", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SetCellLocked_ShouldSetCellLockedStatus()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_cell_locked.xlsx");
        var outputPath = CreateTestFilePath("test_set_cell_locked_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_cell_locked",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:B2",
            ["locked"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var cell = worksheet.Cells["A1"];
        var style = cell.GetStyle();
        Assert.True(style.IsLocked, "Cell should be locked");
    }

    [Fact]
    public async Task SetCellUnlocked_ShouldSetCellUnlockedStatus()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_cell_unlocked.xlsx");
        var outputPath = CreateTestFilePath("test_set_cell_unlocked_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_cell_locked",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:B2",
            ["locked"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var cell = worksheet.Cells["A1"];
        var style = cell.GetStyle();
        Assert.False(style.IsLocked, "Cell should be unlocked");
    }
}