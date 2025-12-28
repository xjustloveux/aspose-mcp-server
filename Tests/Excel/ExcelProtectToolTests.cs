using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelProtectToolTests : ExcelTestBase
{
    private readonly ExcelProtectTool _tool = new();

    #region Protect Tests

    [Fact]
    public async Task Protect_Workbook_ShouldProtectWorkbook()
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
            ["protectWorkbook"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.IsWorkbookProtectedWithPassword);
    }

    [Fact]
    public async Task Protect_Worksheet_ShouldProtectWorksheet()
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
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("worksheet 0", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public async Task Protect_WithInvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_protect_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_protect_invalid_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "protect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["password"] = "test123",
            ["sheetIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task Protect_WithStructureAndWindows_ShouldProtectBoth()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_protect_both.xlsx");
        var outputPath = CreateTestFilePath("test_protect_both_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "protect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["password"] = "test123",
            ["protectWorkbook"] = true,
            ["protectStructure"] = true,
            ["protectWindows"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("workbook", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.IsWorkbookProtectedWithPassword);
    }

    #endregion

    #region Unprotect Tests

    [Fact]
    public async Task Unprotect_Workbook_ShouldUnprotectWorkbook()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_unprotect_workbook.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Protect(ProtectionType.All, "test123");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unprotect_workbook_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unprotect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["password"] = "test123"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("successfully", result);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.IsWorkbookProtectedWithPassword);
    }

    [Fact]
    public async Task Unprotect_Worksheet_ShouldUnprotectWorksheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_unprotect_worksheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unprotect_worksheet_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unprotect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["password"] = "test123",
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("protection removed successfully", result);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public async Task Unprotect_NotProtectedWorksheet_ShouldReturnNotProtectedMessage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_unprotect_not_protected.xlsx");
        var outputPath = CreateTestFilePath("test_unprotect_not_protected_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unprotect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("is not protected", result);
    }

    [Fact]
    public async Task Unprotect_WithInvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_unprotect_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_unprotect_invalid_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unprotect",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Get Protection Tests

    [Fact]
    public async Task GetProtection_AllSheets_ShouldReturnAllSheetsInfo()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_protection_all.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.True(root.TryGetProperty("worksheets", out _));
        Assert.True(root.TryGetProperty("totalWorksheets", out _));
    }

    [Fact]
    public async Task GetProtection_SingleSheet_ShouldReturnSingleSheetInfo()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_protection_single.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
        var worksheets = root.GetProperty("worksheets");
        Assert.True(worksheets[0].GetProperty("isProtected").GetBoolean());
    }

    [Fact]
    public async Task GetProtection_WithInvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_protection_invalid.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task GetProtection_ShouldReturnDetailedProtectionSettings()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_protection_detailed.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var sheet = json.RootElement.GetProperty("worksheets")[0];
        Assert.True(sheet.TryGetProperty("isProtected", out _));
        Assert.True(sheet.TryGetProperty("allowSelectingLockedCell", out _));
        Assert.True(sheet.TryGetProperty("allowFormattingCell", out _));
        Assert.True(sheet.TryGetProperty("allowSorting", out _));
        Assert.True(sheet.TryGetProperty("allowFiltering", out _));
    }

    #endregion

    #region Set Cell Locked Tests

    [Fact]
    public async Task SetCellLocked_ShouldSetCellAsLocked()
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("locked", result);
        using var workbook = new Workbook(outputPath);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.True(style.IsLocked);
    }

    [Fact]
    public async Task SetCellUnlocked_ShouldSetCellAsUnlocked()
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("unlocked", result);
        using var workbook = new Workbook(outputPath);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.False(style.IsLocked);
    }

    [Fact]
    public async Task SetCellLocked_SingleCell_ShouldWork()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_single_cell_locked.xlsx");
        var outputPath = CreateTestFilePath("test_set_single_cell_locked_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_cell_locked",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "C3",
            ["locked"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("locked", result);
        using var workbook = new Workbook(outputPath);
        var style = workbook.Worksheets[0].Cells["C3"].GetStyle();
        Assert.True(style.IsLocked);
    }

    [Fact]
    public async Task SetCellLocked_WithInvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_cell_locked_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_set_cell_locked_invalid_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_cell_locked",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["locked"] = true,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Error Handling Tests

    [Fact]
    public async Task ExecuteAsync_WithUnknownOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_unknown_operation.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid_operation",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task ExecuteAsync_WithMissingPath_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["operation"] = "get"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SetCellLocked_WithMissingRange_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_cell_locked_no_range.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_cell_locked",
            ["path"] = workbookPath,
            ["locked"] = true
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion
}