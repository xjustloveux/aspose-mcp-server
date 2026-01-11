using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelFreezePanesTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void Freeze_ShouldFreezePanesAtSpecifiedPosition()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_output.xlsx");
        var result = _tool.Execute("freeze", workbookPath, row: 2, column: 1, outputPath: outputPath);
        Assert.StartsWith("Frozen panes", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PaneStateType.Frozen, workbook.Worksheets[0].PaneState);
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
    public void Get_WhenFrozen_ShouldReturnFreezeStatus()
    {
        var workbookPath = CreateWorkbookWithFrozenPanes("test_get_frozen.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFrozen").GetBoolean());
    }

    [Fact]
    public void Get_WhenNotFrozen_ShouldReturnNotFrozenStatus()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_not_frozen.xlsx", 10, 5);
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isFrozen").GetBoolean());
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("FREEZE")]
    [InlineData("Freeze")]
    [InlineData("freeze")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 10, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, row: 1, column: 1, outputPath: outputPath);
        Assert.StartsWith("Frozen panes", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

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
