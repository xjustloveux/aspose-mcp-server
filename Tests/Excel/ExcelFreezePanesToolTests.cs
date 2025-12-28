using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFreezePanesToolTests : ExcelTestBase
{
    private readonly ExcelFreezePanesTool _tool = new();

    [Fact]
    public async Task FreezePanes_ShouldFreezePanesAtSpecifiedPosition()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_panes.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_panes_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["row"] = 2,
            ["column"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Frozen panes at row 2, column 1", result);
        Assert.Contains(outputPath, result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, worksheet.PaneState);

        worksheet.GetFreezedPanes(out var frozenRow, out var frozenCol, out _, out _);
        Assert.Equal(3, frozenRow);
        Assert.Equal(2, frozenCol);
    }

    [Fact]
    public async Task FreezePanes_WithSheetIndex_ShouldFreezeCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_sheet_index.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_freeze_sheet_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 1,
            ["row"] = 1,
            ["column"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Frozen panes", result);

        using var workbook = new Workbook(outputPath);
        var worksheet0 = workbook.Worksheets[0];
        var worksheet1 = workbook.Worksheets[1];
        Assert.NotEqual(PaneStateType.Frozen, worksheet0.PaneState);
        Assert.Equal(PaneStateType.Frozen, worksheet1.PaneState);
    }

    [Fact]
    public async Task FreezePanes_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_invalid_sheet.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "freeze",
            ["path"] = workbookPath,
            ["row"] = 1,
            ["column"] = 1,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task UnfreezePanes_ShouldRemoveFreezePanes()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_panes.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.FreezePanes(3, 2, 2, 1);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unfreeze_panes_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unfreeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Unfrozen panes", result);
        Assert.Contains(outputPath, result);

        using var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.NotEqual(PaneStateType.Frozen, resultWorksheet.PaneState);
    }

    [Fact]
    public async Task UnfreezePanes_WhenNotFrozen_ShouldSucceed()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_not_frozen.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_unfreeze_not_frozen_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unfreeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Unfrozen panes", result);
    }

    [Fact]
    public async Task UnfreezePanes_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_invalid_sheet.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "unfreeze",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetFreezePanes_WhenFrozen_ShouldReturnFreezeStatus()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_freeze_status.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.FreezePanes(3, 2, 2, 1);
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
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.True(root.GetProperty("isFrozen").GetBoolean());
        Assert.Equal(2, root.GetProperty("frozenRow").GetInt32());
        Assert.Equal(1, root.GetProperty("frozenColumn").GetInt32());
        Assert.Equal("Panes are frozen", root.GetProperty("status").GetString());
    }

    [Fact]
    public async Task GetFreezePanes_WhenNotFrozen_ShouldReturnNotFrozenStatus()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_not_frozen.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.False(root.GetProperty("isFrozen").GetBoolean());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("frozenRow").ValueKind);
        Assert.Equal(JsonValueKind.Null, root.GetProperty("frozenColumn").ValueKind);
        Assert.Equal("Panes are not frozen", root.GetProperty("status").GetString());
    }

    [Fact]
    public async Task GetFreezePanes_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_invalid_sheet.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_invalid_op.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "invalid",
            ["path"] = workbookPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task FreezePanes_FreezeOnlyRows_ShouldFreezeProperly()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_only_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_only_rows_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["row"] = 3,
            ["column"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Frozen panes at row 3, column 0", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, worksheet.PaneState);
    }

    [Fact]
    public async Task FreezePanes_FreezeOnlyColumns_ShouldFreezeProperly()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_only_cols.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_only_cols_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["row"] = 0,
            ["column"] = 2
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Frozen panes at row 0, column 2", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, worksheet.PaneState);
    }

    [Fact]
    public async Task FreezePanes_WithOutputPath_ShouldNotModifyOriginalFile()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_original.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_output_new.xlsx");
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
        using var originalWorkbook = new Workbook(workbookPath);
        var originalWorksheet = originalWorkbook.Worksheets[0];
        Assert.NotEqual(PaneStateType.Frozen, originalWorksheet.PaneState);

        using var outputWorkbook = new Workbook(outputPath);
        var outputWorksheet = outputWorkbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, outputWorksheet.PaneState);
    }
}