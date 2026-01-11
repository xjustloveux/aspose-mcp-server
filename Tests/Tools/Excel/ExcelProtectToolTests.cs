using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelProtectTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelProtectToolTests : ExcelTestBase
{
    private readonly ExcelProtectTool _tool;

    public ExcelProtectToolTests()
    {
        _tool = new ExcelProtectTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Protect_Worksheet_ShouldProtectWorksheet()
    {
        var workbookPath = CreateExcelWorkbook("test_protect_worksheet.xlsx");
        var outputPath = CreateTestFilePath("test_protect_worksheet_output.xlsx");
        var result = _tool.Execute("protect", workbookPath, sheetIndex: 0, password: "test123", outputPath: outputPath);
        Assert.Contains("protected", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].IsProtected);
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
        var result = _tool.Execute("unprotect", workbookPath, sheetIndex: 0, password: "test123",
            outputPath: outputPath);
        Assert.Contains("protection removed", result);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public void Get_AllSheets_ShouldReturnAllSheetsInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_all.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("worksheets", out _));
    }

    [Fact]
    public void SetCellLocked_ShouldSetCellAsLocked()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_locked.xlsx");
        var outputPath = CreateTestFilePath("test_set_locked_output.xlsx");
        var result = _tool.Execute("set_cell_locked", workbookPath, range: "A1:B2", locked: true,
            outputPath: outputPath);
        Assert.Contains("locked", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells["A1"].GetStyle().IsLocked);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("PROTECT")]
    [InlineData("Protect")]
    [InlineData("protect")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, sheetIndex: 0, password: "test123", outputPath: outputPath);
        Assert.Contains("protected", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Protect_WithSessionId_ShouldProtectInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_protect.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("protect", sessionId: sessionId, sheetIndex: 0, password: "test123");
        Assert.Contains("protected", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].IsProtected);
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
        var result = _tool.Execute("unprotect", sessionId: sessionId, sheetIndex: 0, password: "test123");
        Assert.Contains("protection removed", result);
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(sessionWorkbook.Worksheets[0].IsProtected);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId, sheetIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("worksheets")[0].GetProperty("isProtected").GetBoolean());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateExcelWorkbook("test_session_file.xlsx");
        using (var wb = new Workbook(workbookPath2))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Worksheets[0].Protect(ProtectionType.All, "test123", null);
            wb.Save(workbookPath2);
        }

        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId, sheetIndex: 0);
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}
