using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelSheetToolTests : ExcelTestBase
{
    private readonly ExcelSheetTool _tool;

    public ExcelSheetToolTests()
    {
        _tool = new ExcelSheetTool(SessionManager);
    }

    #region General

    [Fact]
    public void Add_ShouldCreateNewSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        _tool.Execute("add", workbookPath, sheetName: "NewSheet", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var sheetFound = workbook.Worksheets.Any(ws => ws.Name == "NewSheet");
        Assert.True(sheetFound);
    }

    [Fact]
    public void Add_WithInsertAt_ShouldInsertAtPosition()
    {
        var workbookPath = CreateExcelWorkbook("test_add_insert.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_add_insert_output.xlsx");
        _tool.Execute("add", workbookPath, sheetName: "InsertedSheet", insertAt: 0, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var sheetFound = resultWorkbook.Worksheets.Any(ws => ws.Name == "InsertedSheet");
        Assert.True(sheetFound);
    }

    [Fact]
    public void Delete_ShouldDeleteSheet()
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
        else
        {
            // Fallback: verify basic structure in evaluation mode
            var resultWorkbook = new Workbook(outputPath);
            Assert.True(resultWorkbook.Worksheets.Count > 0);
        }
    }

    [Fact]
    public void Get_ShouldReturnAllSheets()
    {
        var workbookPath = CreateExcelWorkbook("test_get.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Save(workbookPath);

        var result = _tool.Execute("get", workbookPath);
        Assert.Contains("count", result);
        Assert.Contains("items", result);
        Assert.Contains("Sheet", result);
    }

    [Fact]
    public void Get_EmptyWorkbook_ShouldReturnValidJson()
    {
        var workbookPath = CreateExcelWorkbook("test_get_minimal.xlsx");
        var result = _tool.Execute("get", workbookPath);
        Assert.Contains("count", result);
        Assert.Contains("items", result);
    }

    [Fact]
    public void Rename_ShouldRenameSheet()
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

    [Fact]
    public void Move_ShouldMoveSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_move.xlsx");
        var workbook = new Workbook(workbookPath);
        var originalFirstSheetName = workbook.Worksheets[0].Name;
        workbook.Worksheets.Add("SheetToMove");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_move_output.xlsx");
        _tool.Execute("move", workbookPath, sheetIndex: 1, insertAt: 0, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        var resultWorkbook = new Workbook(outputPath);
        var sheetFound = resultWorkbook.Worksheets.Any(ws => ws.Name == "SheetToMove");
        Assert.True(sheetFound, "SheetToMove should exist after move");
        Assert.Equal("SheetToMove", resultWorkbook.Worksheets[0].Name);
        Assert.Equal(originalFirstSheetName, resultWorkbook.Worksheets[1].Name);
    }

    [Fact]
    public void Copy_ShouldCopySheet()
    {
        var workbookPath = CreateExcelWorkbook("test_copy.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells[0, 0].Value = "TestData";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_output.xlsx");
        _tool.Execute("copy", workbookPath, sheetIndex: 0, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets.Count >= 2);
        var hasCopiedContent = resultWorkbook.Worksheets
            .Any(ws => ws.Cells[0, 0].Value?.ToString() == "TestData");
        Assert.True(hasCopiedContent);
    }

    [Fact]
    public void Copy_ToExternalFile_ShouldCopySuccessfully()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_external.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells[0, 0].Value = "TestData";
        workbook.Save(workbookPath);

        var copyToPath = CreateTestFilePath("test_copy_external_target.xlsx");
        var result = _tool.Execute("copy", workbookPath, sheetIndex: 0, copyToPath: copyToPath);
        Assert.True(File.Exists(copyToPath));
        Assert.Contains("external file", result);
        var targetWorkbook = new Workbook(copyToPath);
        Assert.Equal("TestData", targetWorkbook.Worksheets[0].Cells[0, 0].Value?.ToString());
    }

    [Fact]
    public void Hide_ShouldHideSheet()
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

    [SkippableFact]
    public void Hide_Toggle_ShouldShowHiddenSheet()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Toggle visibility with additional sheet exceeds limit");
        var workbookPath = CreateExcelWorkbook("test_toggle.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("HiddenSheet");
        workbook.Worksheets["HiddenSheet"].IsVisible = false;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_toggle_output.xlsx");
        var result = _tool.Execute("hide", workbookPath, sheetIndex: 1, outputPath: outputPath);
        Assert.Contains("shown", result);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets["HiddenSheet"].IsVisible);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        _tool.Execute(operation, workbookPath, sheetName: $"Sheet_{operation}", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var sheetFound = workbook.Worksheets.Any(ws => ws.Name == $"Sheet_{operation}");
        Assert.True(sheetFound);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("count", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetName: "Invalid/Name"));
        Assert.Contains("invalid character", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Add_WithNameTooLong_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_long.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetName: new string('A', 32)));
        Assert.Contains("31 characters", ex.Message);
    }

    [Fact]
    public void Add_WithDuplicateName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_dup.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Name = "ExistingSheet";
        workbook.Save(workbookPath);

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetName: "ExistingSheet"));
        Assert.Contains("already exists", ex.Message);
    }

    [SkippableFact]
    public void Delete_LastSheet_ShouldThrowInvalidOperationException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Evaluation warning sheet interferes with last sheet deletion");
        var workbookPath = CreateExcelWorkbook("test_delete_last.xlsx");
        Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("delete", workbookPath, sheetIndex: 0));
    }

    [Fact]
    public void Delete_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, sheetIndex: 999));
    }

    [Fact]
    public void Rename_WithDuplicateName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_rename_dup.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Save(workbookPath);

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("rename", workbookPath, sheetIndex: 0, newName: "Sheet2"));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void Move_ToSamePosition_ShouldReturnNoMoveNeeded()
    {
        var workbookPath = CreateExcelWorkbook("test_move_same.xlsx");
        var result = _tool.Execute("move", workbookPath, sheetIndex: 0, insertAt: 0);
        Assert.Contains("no move needed", result);
    }

    [Fact]
    public void Move_WithoutTargetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_move_no_target.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("move", workbookPath, sheetIndex: 0));
        Assert.Contains("targetIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

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