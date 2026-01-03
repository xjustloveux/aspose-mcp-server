using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelProtectToolTests : ExcelTestBase
{
    private readonly ExcelProtectTool _tool;

    public ExcelProtectToolTests()
    {
        _tool = new ExcelProtectTool(SessionManager);
    }

    #region General Tests

    #region Protect Tests

    [Fact]
    public void Protect_Workbook_ShouldProtectWorkbook()
    {
        var workbookPath = CreateExcelWorkbook("test_protect_workbook.xlsx");
        var outputPath = CreateTestFilePath("test_protect_workbook_output.xlsx");
        var result = _tool.Execute(
            "protect",
            workbookPath,
            password: "test123",
            protectWorkbook: true,
            outputPath: outputPath);
        Assert.Contains("successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.IsWorkbookProtectedWithPassword);
    }

    [Fact]
    public void Protect_Worksheet_ShouldProtectWorksheet()
    {
        var workbookPath = CreateExcelWorkbook("test_protect_worksheet.xlsx");
        var outputPath = CreateTestFilePath("test_protect_worksheet_output.xlsx");
        var result = _tool.Execute(
            "protect",
            workbookPath,
            sheetIndex: 0,
            password: "test123",
            outputPath: outputPath);
        Assert.Contains("worksheet 0", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public void Protect_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_protect_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_protect_invalid_output.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "protect",
            workbookPath,
            sheetIndex: 99,
            password: "test123",
            outputPath: outputPath));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void Protect_WithStructureAndWindows_ShouldProtectBoth()
    {
        var workbookPath = CreateExcelWorkbook("test_protect_both.xlsx");
        var outputPath = CreateTestFilePath("test_protect_both_output.xlsx");
        var result = _tool.Execute(
            "protect",
            workbookPath,
            password: "test123",
            protectWorkbook: true,
            protectStructure: true,
            protectWindows: true,
            outputPath: outputPath);
        Assert.Contains("workbook", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.IsWorkbookProtectedWithPassword);
    }

    #endregion

    #region Unprotect Tests

    [Fact]
    public void Unprotect_Workbook_ShouldUnprotectWorkbook()
    {
        var workbookPath = CreateExcelWorkbook("test_unprotect_workbook.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Protect(ProtectionType.All, "test123");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unprotect_workbook_output.xlsx");
        var result = _tool.Execute(
            "unprotect",
            workbookPath,
            password: "test123",
            outputPath: outputPath);
        Assert.Contains("successfully", result);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.IsWorkbookProtectedWithPassword);
    }

    [Fact]
    public void Unprotect_Worksheet_ShouldUnprotectWorksheet()
    {
        var workbookPath = CreateExcelWorkbook("test_unprotect_worksheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unprotect_worksheet_output.xlsx");
        var result = _tool.Execute(
            "unprotect",
            workbookPath,
            sheetIndex: 0,
            password: "test123",
            outputPath: outputPath);
        Assert.Contains("protection removed successfully", result);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public void Unprotect_NotProtectedWorksheet_ShouldReturnNotProtectedMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_unprotect_not_protected.xlsx");
        var outputPath = CreateTestFilePath("test_unprotect_not_protected_output.xlsx");
        var result = _tool.Execute(
            "unprotect",
            workbookPath,
            sheetIndex: 0,
            outputPath: outputPath);
        Assert.Contains("is not protected", result);
    }

    [Fact]
    public void Unprotect_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_unprotect_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_unprotect_invalid_output.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "unprotect",
            workbookPath,
            sheetIndex: 99,
            outputPath: outputPath));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Get Protection Tests

    [Fact]
    public void GetProtection_AllSheets_ShouldReturnAllSheetsInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_protection_all.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.True(root.TryGetProperty("worksheets", out _));
        Assert.True(root.TryGetProperty("totalWorksheets", out _));
    }

    [Fact]
    public void GetProtection_SingleSheet_ShouldReturnSingleSheetInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_protection_single.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath,
            sheetIndex: 0);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
        var worksheets = root.GetProperty("worksheets");
        Assert.True(worksheets[0].GetProperty("isProtected").GetBoolean());
    }

    [Fact]
    public void GetProtection_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_protection_invalid.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "get",
            workbookPath,
            sheetIndex: 99));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void GetProtection_ShouldReturnDetailedProtectionSettings()
    {
        var workbookPath = CreateExcelWorkbook("test_get_protection_detailed.xlsx");
        var result = _tool.Execute(
            "get",
            workbookPath,
            sheetIndex: 0);
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
    public void SetCellLocked_ShouldSetCellAsLocked()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_cell_locked.xlsx");
        var outputPath = CreateTestFilePath("test_set_cell_locked_output.xlsx");
        var result = _tool.Execute(
            "set_cell_locked",
            workbookPath,
            range: "A1:B2",
            locked: true,
            outputPath: outputPath);
        Assert.Contains("locked", result);
        using var workbook = new Workbook(outputPath);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.True(style.IsLocked);
    }

    [Fact]
    public void SetCellUnlocked_ShouldSetCellAsUnlocked()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_cell_unlocked.xlsx");
        var outputPath = CreateTestFilePath("test_set_cell_unlocked_output.xlsx");
        var result = _tool.Execute(
            "set_cell_locked",
            workbookPath,
            range: "A1:B2",
            locked: false,
            outputPath: outputPath);
        Assert.Contains("unlocked", result);
        using var workbook = new Workbook(outputPath);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.False(style.IsLocked);
    }

    [Fact]
    public void SetCellLocked_SingleCell_ShouldWork()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_single_cell_locked.xlsx");
        var outputPath = CreateTestFilePath("test_set_single_cell_locked_output.xlsx");
        var result = _tool.Execute(
            "set_cell_locked",
            workbookPath,
            range: "C3",
            locked: true,
            outputPath: outputPath);
        Assert.Contains("locked", result);
        using var workbook = new Workbook(outputPath);
        var style = workbook.Worksheets[0].Cells["C3"].GetStyle();
        Assert.True(style.IsLocked);
    }

    [Fact]
    public void SetCellLocked_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_cell_locked_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_set_cell_locked_invalid_output.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_cell_locked",
            workbookPath,
            sheetIndex: 99,
            range: "A1",
            locked: true,
            outputPath: outputPath));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Error Handling Tests

    [Fact]
    public void ExecuteAsync_WithMissingPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "get",
            ""));
    }

    [Fact]
    public void SetCellLocked_WithMissingRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_cell_locked_no_range.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_cell_locked",
            workbookPath,
            range: "",
            locked: true));
    }

    #endregion

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_operation.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid_operation",
            workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void Protect_WithMissingPassword_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_protect_no_password.xlsx");
        var outputPath = CreateTestFilePath("test_protect_no_password_output.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "protect",
            workbookPath,
            protectWorkbook: true,
            password: "",
            outputPath: outputPath));
        Assert.Contains("password", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetProtection_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_protection.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "get",
            sessionId: sessionId,
            sheetIndex: 0);
        var json = JsonDocument.Parse(result);
        var worksheets = json.RootElement.GetProperty("worksheets");
        Assert.True(worksheets[0].GetProperty("isProtected").GetBoolean());
    }

    [Fact]
    public void Protect_WithSessionId_ShouldProtectInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_protect.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "protect",
            sessionId: sessionId,
            sheetIndex: 0,
            password: "test123");
        Assert.Contains("worksheet 0", result);

        // Verify in-memory workbook has protection
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public void SetCellLocked_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_set_locked.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "set_cell_locked",
            sessionId: sessionId,
            range: "A1:B2",
            locked: true);
        Assert.Contains("locked", result);

        // Verify in-memory workbook has locked cells
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.True(style.IsLocked);
    }

    [Fact]
    public void Unprotect_WithSessionId_ShouldUnprotectInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_unprotect.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "unprotect",
            sessionId: sessionId,
            sheetIndex: 0,
            password: "test123");
        Assert.Contains("protection removed successfully", result);

        // Verify in-memory workbook is unprotected
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(sessionWorkbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute(
            "get",
            sessionId: "invalid_session_id"));
    }

    #endregion
}