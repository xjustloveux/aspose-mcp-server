using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelFreezePanesToolTests : ExcelTestBase
{
    private readonly ExcelFreezePanesTool _tool;

    public ExcelFreezePanesToolTests()
    {
        _tool = new ExcelFreezePanesTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void FreezePanes_ShouldFreezePanesAtSpecifiedPosition()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_panes.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_panes_output.xlsx");
        var result = _tool.Execute(
            "freeze",
            workbookPath,
            row: 2,
            column: 1,
            outputPath: outputPath);
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
    public void FreezePanes_WithSheetIndex_ShouldFreezeCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_sheet_index.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_freeze_sheet_index_output.xlsx");
        var result = _tool.Execute(
            "freeze",
            workbookPath,
            sheetIndex: 1,
            row: 1,
            column: 1,
            outputPath: outputPath);
        Assert.Contains("Frozen panes", result);

        using var workbook = new Workbook(outputPath);
        var worksheet0 = workbook.Worksheets[0];
        var worksheet1 = workbook.Worksheets[1];
        Assert.NotEqual(PaneStateType.Frozen, worksheet0.PaneState);
        Assert.Equal(PaneStateType.Frozen, worksheet1.PaneState);
    }

    [Fact]
    public void UnfreezePanes_ShouldRemoveFreezePanes()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_panes.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.FreezePanes(3, 2, 2, 1);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unfreeze_panes_output.xlsx");
        var result = _tool.Execute(
            "unfreeze",
            workbookPath,
            outputPath: outputPath);
        Assert.Contains("Unfrozen panes", result);
        Assert.Contains(outputPath, result);

        using var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.NotEqual(PaneStateType.Frozen, resultWorksheet.PaneState);
    }

    [Fact]
    public void UnfreezePanes_WhenNotFrozen_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_not_frozen.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_unfreeze_not_frozen_output.xlsx");
        var result = _tool.Execute(
            "unfreeze",
            workbookPath,
            outputPath: outputPath);
        Assert.Contains("Unfrozen panes", result);
    }

    [Fact]
    public void GetFreezePanes_WhenFrozen_ShouldReturnFreezeStatus()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_freeze_status.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.FreezePanes(3, 2, 2, 1);
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.True(root.GetProperty("isFrozen").GetBoolean());
        Assert.Equal(2, root.GetProperty("frozenRow").GetInt32());
        Assert.Equal(1, root.GetProperty("frozenColumn").GetInt32());
        Assert.Equal("Panes are frozen", root.GetProperty("status").GetString());
    }

    [Fact]
    public void GetFreezePanes_WhenNotFrozen_ShouldReturnNotFrozenStatus()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_not_frozen.xlsx", 10, 5);
        var result = _tool.Execute(
            "get",
            workbookPath);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.False(root.GetProperty("isFrozen").GetBoolean());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("frozenRow").ValueKind);
        Assert.Equal(JsonValueKind.Null, root.GetProperty("frozenColumn").ValueKind);
        Assert.Equal("Panes are not frozen", root.GetProperty("status").GetString());
    }

    [Fact]
    public void FreezePanes_FreezeOnlyRows_ShouldFreezeProperly()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_only_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_only_rows_output.xlsx");
        var result = _tool.Execute(
            "freeze",
            workbookPath,
            row: 3,
            column: 0,
            outputPath: outputPath);
        Assert.Contains("Frozen panes at row 3, column 0", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, worksheet.PaneState);
    }

    [Fact]
    public void FreezePanes_FreezeOnlyColumns_ShouldFreezeProperly()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_only_cols.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_only_cols_output.xlsx");
        var result = _tool.Execute(
            "freeze",
            workbookPath,
            row: 0,
            column: 2,
            outputPath: outputPath);
        Assert.Contains("Frozen panes at row 0, column 2", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, worksheet.PaneState);
    }

    [Fact]
    public void FreezePanes_WithOutputPath_ShouldNotModifyOriginalFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_original.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_output_new.xlsx");
        _tool.Execute(
            "freeze",
            workbookPath,
            row: 1,
            column: 1,
            outputPath: outputPath);
        using var originalWorkbook = new Workbook(workbookPath);
        var originalWorksheet = originalWorkbook.Worksheets[0];
        Assert.NotEqual(PaneStateType.Frozen, originalWorksheet.PaneState);

        using var outputWorkbook = new Workbook(outputPath);
        var outputWorksheet = outputWorkbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, outputWorksheet.PaneState);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void FreezePanes_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_invalid_sheet.xlsx", 10, 5);
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "freeze",
            workbookPath,
            sheetIndex: 99,
            row: 1,
            column: 1));
    }

    [Fact]
    public void UnfreezePanes_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_invalid_sheet.xlsx", 10, 5);
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "unfreeze",
            workbookPath,
            sheetIndex: 99));
    }

    [Fact]
    public void GetFreezePanes_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_invalid_sheet.xlsx", 10, 5);
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "get",
            workbookPath,
            sheetIndex: 99));
    }

    [Fact]
    public void ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_invalid_op.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid",
            workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Freeze_WithDefaultRow_ShouldFreezeWithRowZero()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_default_row.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_default_row_output.xlsx");

        // Act - row defaults to 0, which is valid (freeze only columns)
        var result = _tool.Execute(
            "freeze",
            workbookPath,
            column: 1,
            outputPath: outputPath);
        Assert.Contains("Frozen panes at row 0, column 1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.FreezePanes(3, 2, 2, 1);
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "get",
            sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFrozen").GetBoolean());
    }

    [Fact]
    public void Freeze_WithSessionId_ShouldFreezeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_freeze.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "freeze",
            sessionId: sessionId,
            row: 2,
            column: 1);
        Assert.Contains("Frozen panes", result);

        // Verify in-memory workbook has frozen panes
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Unfreeze_WithSessionId_ShouldUnfreezeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_unfreeze.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.FreezePanes(3, 2, 2, 1);
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "unfreeze",
            sessionId: sessionId);
        Assert.Contains("Unfrozen panes", result);

        // Verify in-memory workbook has no frozen panes
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotEqual(PaneStateType.Frozen, sessionWorkbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Freeze_WithSessionId_ShouldNotModifyOriginalFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_freeze_original.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute(
            "freeze",
            sessionId: sessionId,
            row: 1,
            column: 1);

        // Assert - original file should not have frozen panes
        using var originalWorkbook = new Workbook(workbookPath);
        Assert.NotEqual(PaneStateType.Frozen, originalWorkbook.Worksheets[0].PaneState);

        // But session workbook should have frozen panes
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(PaneStateType.Frozen, sessionWorkbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}