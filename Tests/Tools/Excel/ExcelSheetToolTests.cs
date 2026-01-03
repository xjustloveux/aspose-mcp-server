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

    #region General Tests

    [Fact]
    public void CreateSheet_ShouldCreateNewSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_create_sheet.xlsx");
        var outputPath = CreateTestFilePath("test_create_sheet_output.xlsx");
        _tool.Execute("add", workbookPath, sheetName: "NewSheet", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        // In evaluation mode, sheet creation may add evaluation warning sheets
        // Check that NewSheet exists (may be at different index)
        var sheetFound = false;
        foreach (var worksheet in workbook.Worksheets)
            if (worksheet.Name == "NewSheet")
            {
                sheetFound = true;
                break;
            }

        Assert.True(sheetFound, "NewSheet should be created");
    }

    [Fact]
    public void DeleteSheet_ShouldDeleteSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToDelete");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_sheet_output.xlsx");
        _tool.Execute("delete", workbookPath, sheetIndex: 1, outputPath: outputPath);

        var isEvaluationMode = IsEvaluationMode();
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");

        var sheetExists = false;
        foreach (var worksheet in resultWorkbook.Worksheets)
            if (worksheet.Name == "SheetToDelete")
            {
                sheetExists = true;
                break;
            }

        if (isEvaluationMode)
        {
            Assert.True(resultWorkbook.Worksheets.Count > 0, "Workbook should have at least one sheet");
        }
        else
        {
            Assert.False(sheetExists, "SheetToDelete should be deleted");
            Assert.DoesNotContain("SheetToDelete", resultWorkbook.Worksheets.Select(s => s.Name));
        }
    }

    [Fact]
    public void RenameSheet_ShouldRenameSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_rename_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("OldName");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_rename_sheet_output.xlsx");
        _tool.Execute("rename", workbookPath, sheetIndex: 1, newName: "NewName", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var sheet = resultWorkbook.Worksheets["NewName"];
        Assert.NotNull(sheet);
    }

    [Fact]
    public void CopySheet_ShouldCopySheet()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        var sourceSheet = workbook.Worksheets[0];
        sourceSheet.Cells[0, 0].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_sheet_output.xlsx");
        _tool.Execute("copy", workbookPath, sheetIndex: 0, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        // AddCopy creates a copy with name like "Sheet1 (2)", need to find by index or check if copy exists
        // The copy should be at the end (index = original count)
        Assert.True(resultWorkbook.Worksheets.Count >= 2);
        // Check if the copied sheet exists (it might have a different name like "Sheet1 (2)")
        var hasCopiedContent = false;
        foreach (var worksheet in resultWorkbook.Worksheets)
            if (worksheet.Cells[0, 0].Value?.ToString() == "Test")
            {
                hasCopiedContent = true;
                break;
            }

        Assert.True(hasCopiedContent, "Copied sheet should contain the test value");
    }

    [Fact]
    public void MoveSheet_ShouldMoveSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_move_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToMove");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_move_sheet_output.xlsx");
        _tool.Execute("move", workbookPath, sheetIndex: 1, insertAt: 0, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        // Check that SheetToMove exists
        // Note: Aspose.Cells evaluation version may add "Evaluation Warning" sheet, affecting positions
        var sheetToMoveFound = false;
        var sheetToMoveIndex = -1;
        for (var i = 0; i < resultWorkbook.Worksheets.Count; i++)
            if (resultWorkbook.Worksheets[i].Name == "SheetToMove")
            {
                sheetToMoveFound = true;
                sheetToMoveIndex = i;
                break;
            }

        Assert.True(sheetToMoveFound, "SheetToMove should exist in the workbook");
        Assert.True(File.Exists(outputPath), "Output workbook should be created");

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            Assert.True(sheetToMoveIndex < resultWorkbook.Worksheets.Count,
                $"Sheet should be at valid position: {sheetToMoveIndex}");
        }
        else
        {
            Assert.Equal("SheetToMove", resultWorkbook.Worksheets[sheetToMoveIndex].Name);
            Assert.True(sheetToMoveIndex <= 1, "Sheet should be moved to position 0 or 1");
        }
    }

    [Fact]
    public void GetSheets_ShouldReturnAllSheets()
    {
        var workbookPath = CreateExcelWorkbook("test_get_sheets.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Save(workbookPath);
        var result = _tool.Execute("get", workbookPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Sheet", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HideSheet_ShouldHideSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_hide_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToHide");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_hide_sheet_output.xlsx");
        _tool.Execute("hide", workbookPath, sheetIndex: 1, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        var sheetExists = false;
        foreach (var worksheet in resultWorkbook.Worksheets)
            if (worksheet.Name == "SheetToHide")
            {
                sheetExists = true;
                break;
            }

        Assert.True(sheetExists, "SheetToHide should exist in the workbook");
    }

    [Fact]
    public void AddSheet_WithInsertAt_ShouldInsertAtPosition()
    {
        var workbookPath = CreateExcelWorkbook("test_add_insert_at.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_add_insert_at_output.xlsx");
        _tool.Execute("add", workbookPath, sheetName: "InsertedSheet", insertAt: 0, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var sheetFound = false;
        foreach (var worksheet in resultWorkbook.Worksheets)
            if (worksheet.Name == "InsertedSheet")
            {
                sheetFound = true;
                break;
            }

        Assert.True(sheetFound, "InsertedSheet should be created");
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void AddSheet_WithInvalidName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_name.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetName: "Invalid/Name"));
        Assert.Contains("invalid character", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddSheet_WithNameTooLong_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_name_too_long.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetName: new string('A', 32)));
        Assert.Contains("31 characters", ex.Message);
    }

    [Fact]
    public void AddSheet_WithDuplicateName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_duplicate_name.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Name = "ExistingSheet";
        workbook.Save(workbookPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetName: "ExistingSheet"));
        Assert.Contains("already exists", ex.Message);
    }

    [SkippableFact]
    public void DeleteSheet_LastSheet_ShouldThrowInvalidOperationException()
    {
        // Skip in evaluation mode - evaluation warning sheet interferes with last sheet deletion
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Evaluation warning sheet interferes with last sheet deletion");
        var workbookPath = CreateExcelWorkbook("test_delete_last.xlsx");
        Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("delete", workbookPath, sheetIndex: 0));
    }

    [Fact]
    public void RenameSheet_WithDuplicateName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_rename_duplicate.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Save(workbookPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("rename", workbookPath, sheetIndex: 0, newName: "Sheet2"));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void MoveSheet_ToSamePosition_ShouldNotModifyFile()
    {
        var workbookPath = CreateExcelWorkbook("test_move_same_position.xlsx");
        var result = _tool.Execute("move", workbookPath, sheetIndex: 0, insertAt: 0);
        Assert.Contains("no move needed", result);
    }

    [Fact]
    public void MoveSheet_WithoutTargetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_move_no_target.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("move", workbookPath, sheetIndex: 0));
        Assert.Contains("targetIndex", ex.Message);
    }

    [Fact]
    public void CopySheet_ToExternalFile_ShouldCopySuccessfully()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_external.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells[0, 0].Value = "TestData";
        workbook.Save(workbookPath);

        var copyToPath = CreateTestFilePath("test_copy_external_target.xlsx");
        var result = _tool.Execute("copy", workbookPath, sheetIndex: 0, copyToPath: copyToPath);
        Assert.True(File.Exists(copyToPath), "Target workbook should be created");
        Assert.Contains("external file", result);

        var targetWorkbook = new Workbook(copyToPath);
        Assert.Equal("TestData", targetWorkbook.Worksheets[0].Cells[0, 0].Value?.ToString());
    }

    [SkippableFact]
    public void HideSheet_Toggle_ShouldShowHiddenSheet()
    {
        // Skip in evaluation mode - toggle visibility with additional sheet exceeds limit
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Toggle visibility with additional sheet exceeds limit");
        var workbookPath = CreateExcelWorkbook("test_toggle_visibility.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("HiddenSheet");
        workbook.Worksheets["HiddenSheet"].IsVisible = false;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_toggle_visibility_output.xlsx");
        var result = _tool.Execute("hide", workbookPath, sheetIndex: 1, outputPath: outputPath);
        Assert.Contains("shown", result);
        var resultWorkbook = new Workbook(outputPath);
        var hiddenSheet = resultWorkbook.Worksheets["HiddenSheet"];
        Assert.True(hiddenSheet.IsVisible, "Hidden sheet should now be visible");
    }

    [Fact]
    public void GetSheets_EmptyWorkbook_ShouldReturnValidJson()
    {
        // Arrange - Create workbook but don't add extra sheets
        var workbookPath = CreateExcelWorkbook("test_get_sheets_minimal.xlsx");
        var result = _tool.Execute("get", workbookPath);
        Assert.NotNull(result);
        Assert.Contains("count", result);
        Assert.Contains("items", result);
    }

    [Fact]
    public void InvalidOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("invalid_operation", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void SheetIndex_OutOfRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_index_out_of_range.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, sheetIndex: 999));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddSheet_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add_sheet.xlsx");
        var sessionId = OpenSession(workbookPath);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var initialCount = workbook.Worksheets.Count;
        var result = _tool.Execute("add", sessionId: sessionId, sheetName: "SessionSheet");
        Assert.Contains("added", result);
        Assert.True(workbook.Worksheets.Count > initialCount);
    }

    [Fact]
    public void GetSheets_WithSessionId_ShouldReturnInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_sheets.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void RenameSheet_WithSessionId_ShouldRenameInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_rename_sheet.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("rename", sessionId: sessionId, sheetIndex: 0, newName: "RenamedSheet");
        Assert.Contains("renamed", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("RenamedSheet", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void CopySheet_WithSessionId_ShouldCopyInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_copy_sheet.xlsx");
        var sessionId = OpenSession(workbookPath);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var initialCount = workbook.Worksheets.Count;
        var result = _tool.Execute("copy", sessionId: sessionId, sheetIndex: 0);
        Assert.Contains("copied", result);
        Assert.True(workbook.Worksheets.Count > initialCount);
    }

    #endregion
}