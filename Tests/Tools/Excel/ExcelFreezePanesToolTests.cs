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

    private string CreateWorkbookWithFrozenPanes(string fileName, int frozenRow = 2, int frozenCol = 1)
    {
        var path = CreateExcelWorkbookWithData(fileName, 10, 5);
        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[0];
        worksheet.FreezePanes(frozenRow + 1, frozenCol + 1, frozenRow, frozenCol);
        workbook.Save(path);
        return path;
    }

    #region General

    [Fact]
    public void Freeze_ShouldFreezePanesAtSpecifiedPosition()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_output.xlsx");
        var result = _tool.Execute("freeze", workbookPath, row: 2, column: 1, outputPath: outputPath);
        Assert.StartsWith("Frozen panes", result);
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(PaneStateType.Frozen, worksheet.PaneState);
        worksheet.GetFreezedPanes(out var frozenRow, out var frozenCol, out _, out _);
        Assert.Equal(3, frozenRow);
        Assert.Equal(2, frozenCol);
    }

    [Fact]
    public void Freeze_WithSheetIndex_ShouldFreezeCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_sheet.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_freeze_sheet_output.xlsx");
        var result = _tool.Execute("freeze", workbookPath, sheetIndex: 1, row: 1, column: 1, outputPath: outputPath);
        Assert.StartsWith("Frozen panes", result);
        using var workbook = new Workbook(outputPath);
        Assert.NotEqual(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[1].PaneState);
    }

    [Fact]
    public void Freeze_OnlyRows_ShouldFreezeProperly()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_rows_output.xlsx");
        var result = _tool.Execute("freeze", workbookPath, row: 3, column: 0, outputPath: outputPath);
        Assert.StartsWith("Frozen panes", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Freeze_OnlyColumns_ShouldFreezeProperly()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_cols.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_cols_output.xlsx");
        var result = _tool.Execute("freeze", workbookPath, row: 0, column: 2, outputPath: outputPath);
        Assert.StartsWith("Frozen panes", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Freeze_WithOutputPath_ShouldNotModifyOriginalFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_original.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_output_new.xlsx");
        _tool.Execute("freeze", workbookPath, row: 1, column: 1, outputPath: outputPath);
        using var originalWorkbook = new Workbook(workbookPath);
        Assert.NotEqual(PaneStateType.Frozen, originalWorkbook.Worksheets[0].PaneState);
        using var outputWorkbook = new Workbook(outputPath);
        Assert.Equal(PaneStateType.Frozen, outputWorkbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Unfreeze_ShouldRemoveFreezePanes()
    {
        var workbookPath = CreateWorkbookWithFrozenPanes("test_unfreeze.xlsx");
        var outputPath = CreateTestFilePath("test_unfreeze_output.xlsx");
        var result = _tool.Execute("unfreeze", workbookPath, outputPath: outputPath);
        Assert.StartsWith("Unfrozen panes", result);
        using var workbook = new Workbook(outputPath);
        Assert.NotEqual(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Unfreeze_WhenNotFrozen_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_not_frozen.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_unfreeze_not_frozen_output.xlsx");
        var result = _tool.Execute("unfreeze", workbookPath, outputPath: outputPath);
        Assert.StartsWith("Unfrozen panes", result);
    }

    [Fact]
    public void Unfreeze_WithSheetIndex_ShouldUnfreezeCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_sheet.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[0].FreezePanes(3, 2, 2, 1);
            wb.Worksheets[1].FreezePanes(3, 2, 2, 1);
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unfreeze_sheet_output.xlsx");
        _tool.Execute("unfreeze", workbookPath, sheetIndex: 1, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
        Assert.NotEqual(PaneStateType.Frozen, workbook.Worksheets[1].PaneState);
    }

    [Fact]
    public void Get_WhenFrozen_ShouldReturnFreezeStatus()
    {
        var workbookPath = CreateWorkbookWithFrozenPanes("test_get_frozen.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.True(root.GetProperty("isFrozen").GetBoolean());
        Assert.Equal(2, root.GetProperty("frozenRow").GetInt32());
        Assert.Equal(1, root.GetProperty("frozenColumn").GetInt32());
        Assert.Equal("Panes are frozen", root.GetProperty("status").GetString());
    }

    [Fact]
    public void Get_WhenNotFrozen_ShouldReturnNotFrozenStatus()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_not_frozen.xlsx", 10, 5);
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.False(root.GetProperty("isFrozen").GetBoolean());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("frozenRow").ValueKind);
        Assert.Equal(JsonValueKind.Null, root.GetProperty("frozenColumn").ValueKind);
        Assert.Equal("Panes are not frozen", root.GetProperty("status").GetString());
    }

    [Fact]
    public void Get_ShouldReturnWorksheetName()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_worksheet.xlsx", 10, 5);
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
    }

    [Fact]
    public void Get_WithSheetIndex_ShouldReturnCorrectSheetStatus()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_sheet.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].FreezePanes(3, 2, 2, 1);
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath, sheetIndex: 1);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFrozen").GetBoolean());
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
    }

    [Theory]
    [InlineData("FREEZE")]
    [InlineData("Freeze")]
    [InlineData("freeze")]
    public void Operation_ShouldBeCaseInsensitive_Freeze(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 10, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, row: 1, column: 1, outputPath: outputPath);
        Assert.StartsWith("Frozen panes", result);
    }

    [Theory]
    [InlineData("UNFREEZE")]
    [InlineData("Unfreeze")]
    [InlineData("unfreeze")]
    public void Operation_ShouldBeCaseInsensitive_Unfreeze(string operation)
    {
        var workbookPath = CreateWorkbookWithFrozenPanes($"test_case_unfreeze_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_unfreeze_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, outputPath: outputPath);
        Assert.StartsWith("Unfrozen panes", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_get_{operation}.xlsx", 10, 5);
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("isFrozen", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Freeze_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_invalid_sheet.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("freeze", workbookPath, sheetIndex: 99, row: 1, column: 1));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Unfreeze_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze_invalid_sheet.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unfreeze", workbookPath, sheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_invalid_sheet.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, sheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Freeze_WithSessionId_ShouldFreezeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_freeze.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("freeze", sessionId: sessionId, row: 2, column: 1);
        Assert.StartsWith("Frozen panes", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Unfreeze_WithSessionId_ShouldUnfreezeInMemory()
    {
        var workbookPath = CreateWorkbookWithFrozenPanes("test_session_unfreeze.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("unfreeze", sessionId: sessionId);
        Assert.StartsWith("Unfrozen panes", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotEqual(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithFrozenPanes("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFrozen").GetBoolean());
    }

    [Fact]
    public void Freeze_WithSessionId_ShouldNotModifyOriginalFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_freeze_original.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("freeze", sessionId: sessionId, row: 1, column: 1);
        using var originalWorkbook = new Workbook(workbookPath);
        Assert.NotEqual(PaneStateType.Frozen, originalWorkbook.Worksheets[0].PaneState);
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(PaneStateType.Frozen, sessionWorkbook.Worksheets[0].PaneState);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbookWithData("test_path_file.xlsx", 10, 5);
        var sessionWorkbook = CreateWorkbookWithFrozenPanes("test_session_file.xlsx");
        using (var wb = new Workbook(sessionWorkbook))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Save(sessionWorkbook);
        }

        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}