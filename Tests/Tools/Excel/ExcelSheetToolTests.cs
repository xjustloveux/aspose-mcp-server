using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelSheetTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelSheetToolTests : ExcelTestBase
{
    private readonly ExcelSheetTool _tool;

    public ExcelSheetToolTests()
    {
        _tool = new ExcelSheetTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldCreateNewSheetAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        _tool.Execute("add", workbookPath, sheetName: "NewSheet", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var sheetFound = workbook.Worksheets.Any(ws => ws.Name == "NewSheet");
        Assert.True(sheetFound);
    }

    [Fact]
    public void Delete_ShouldDeleteSheetAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_delete.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToDelete");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        _tool.Execute("delete", workbookPath, sheetIndex: 1, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));

        if (!IsEvaluationMode())
        {
            var resultWorkbook = new Workbook(outputPath);
            Assert.DoesNotContain("SheetToDelete", resultWorkbook.Worksheets.Select(s => s.Name));
        }
    }

    [Fact]
    public void Get_ShouldReturnAllSheetsFromFile()
    {
        var workbookPath = CreateExcelWorkbook("test_get.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Save(workbookPath);

        var result = _tool.Execute("get", workbookPath);
        Assert.Contains("count", result);
        Assert.Contains("items", result);
    }

    [Fact]
    public void Rename_ShouldRenameSheetAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_rename.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("OldName");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_rename_output.xlsx");
        _tool.Execute("rename", workbookPath, sheetIndex: 1, newName: "NewName", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        Assert.NotNull(resultWorkbook.Worksheets["NewName"]);
    }

    [SkippableFact]
    public void Move_ShouldMoveSheetAndPersistToFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells,
            "Evaluation mode inserts watermark sheet affecting move position");
        var workbookPath = CreateExcelWorkbook("test_move.xlsx");
        var workbook = new Workbook(workbookPath);
        var originalFirstSheetName = workbook.Worksheets[0].Name;
        workbook.Worksheets.Add("SheetToMove");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_move_output.xlsx");
        _tool.Execute("move", workbookPath, sheetIndex: 1, insertAt: 0, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("SheetToMove", resultWorkbook.Worksheets[0].Name);
        Assert.Equal(originalFirstSheetName, resultWorkbook.Worksheets[1].Name);
    }

    [Fact]
    public void Copy_ShouldCopySheetAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_copy.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells[0, 0].Value = "TestData";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_output.xlsx");
        _tool.Execute("copy", workbookPath, sheetIndex: 0, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets.Count >= 2);
    }

    [Fact]
    public void Hide_ShouldHideSheetAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_hide.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToHide");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_hide_output.xlsx");
        var result = _tool.Execute("hide", workbookPath, sheetIndex: 1, outputPath: outputPath);
        Assert.Contains("hidden", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        _tool.Execute(operation, workbookPath, sheetName: $"Sheet_{operation}", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var sheetFound = workbook.Worksheets.Any(ws => ws.Name == $"Sheet_{operation}");
        Assert.True(sheetFound);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var initialCount = workbook.Worksheets.Count;
        var result = _tool.Execute("add", sessionId: sessionId, sheetName: "SessionSheet");
        Assert.Contains("added", result);
        Assert.Contains("session", result);
        Assert.True(workbook.Worksheets.Count > initialCount);
    }

    [Fact]
    public void Get_WithSessionId_ShouldReturnInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("count", result);
    }

    [Fact]
    public void Rename_WithSessionId_ShouldRenameInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_rename.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("rename", sessionId: sessionId, sheetIndex: 0, newName: "RenamedSheet");
        Assert.Contains("renamed", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("RenamedSheet", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void Copy_WithSessionId_ShouldCopyInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_copy.xlsx");
        var sessionId = OpenSession(workbookPath);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var initialCount = workbook.Worksheets.Count;
        var result = _tool.Execute("copy", sessionId: sessionId, sheetIndex: 0);
        Assert.Contains("copied", result);
        Assert.Contains("session", result);
        Assert.True(workbook.Worksheets.Count > initialCount);
    }

    [SkippableFact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Evaluation mode adds extra watermark worksheets");
        var workbookPath = CreateExcelWorkbook("test_session_delete.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToDelete");
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        var wb = SessionManager.GetDocument<Workbook>(sessionId);
        var initialCount = wb.Worksheets.Count;
        var result = _tool.Execute("delete", sessionId: sessionId, sheetIndex: 1);
        Assert.Contains("deleted", result);
        Assert.Contains("session", result);
        Assert.True(wb.Worksheets.Count < initialCount);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateExcelWorkbook("test_session_file.xlsx");
        var workbook = new Workbook(workbookPath2);
        workbook.Worksheets[0].Name = "SessionSheet";
        workbook.Save(workbookPath2);

        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}
