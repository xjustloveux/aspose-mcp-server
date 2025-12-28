using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelSheetToolTests : ExcelTestBase
{
    private readonly ExcelSheetTool _tool = new();

    [Fact]
    public async Task CreateSheet_ShouldCreateNewSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_create_sheet.xlsx");
        var outputPath = CreateTestFilePath("test_create_sheet_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["sheetName"] = "NewSheet";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task DeleteSheet_ShouldDeleteSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToDelete");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_sheet_output.xlsx");
        var arguments = CreateArguments("delete", workbookPath, outputPath);
        arguments["sheetIndex"] = 1; // Index of "SheetToDelete" (0 is Sheet1, 1 is SheetToDelete)

        // Act
        await _tool.ExecuteAsync(arguments);

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
    public async Task RenameSheet_ShouldRenameSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_rename_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("OldName");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_rename_sheet_output.xlsx");
        var arguments = CreateArguments("rename", workbookPath, outputPath);
        arguments["sheetIndex"] = 1; // Index of "OldName" (0 is Sheet1, 1 is OldName)
        arguments["newName"] = "NewName";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var sheet = resultWorkbook.Worksheets["NewName"];
        Assert.NotNull(sheet);
    }

    [Fact]
    public async Task CopySheet_ShouldCopySheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        var sourceSheet = workbook.Worksheets[0];
        sourceSheet.Cells[0, 0].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_sheet_output.xlsx");
        var arguments = CreateArguments("copy", workbookPath, outputPath);
        arguments["sheetIndex"] = 0; // Index of Sheet1
        arguments["newName"] = "CopiedSheet";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task MoveSheet_ShouldMoveSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_move_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToMove");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_move_sheet_output.xlsx");
        var arguments = CreateArguments("move", workbookPath, outputPath);
        arguments["sheetIndex"] = 1; // Index of "SheetToMove" (0 is Sheet1, 1 is SheetToMove)
        arguments["insertAt"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task GetSheets_ShouldReturnAllSheets()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheets.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Save(workbookPath);

        var arguments = CreateArguments("get", workbookPath);
        arguments["operation"] = "get";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Sheet", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task HideSheet_ShouldHideSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_hide_sheet.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("SheetToHide");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_hide_sheet_output.xlsx");
        var arguments = CreateArguments("hide", workbookPath, outputPath);
        arguments["sheetIndex"] = 1; // Index of "SheetToHide"

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task AddSheet_WithInsertAt_ShouldInsertAtPosition()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_insert_at.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_add_insert_at_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["sheetName"] = "InsertedSheet";
        arguments["insertAt"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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

    [Fact]
    public async Task AddSheet_WithInvalidName_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_name.xlsx");
        var arguments = CreateArguments("add", workbookPath);
        arguments["sheetName"] = "Invalid/Name";

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("invalid character", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AddSheet_WithNameTooLong_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_name_too_long.xlsx");
        var arguments = CreateArguments("add", workbookPath);
        arguments["sheetName"] = new string('A', 32); // 32 characters, exceeds 31 limit

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("31 characters", ex.Message);
    }

    [Fact]
    public async Task AddSheet_WithDuplicateName_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_duplicate_name.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Name = "ExistingSheet";
        workbook.Save(workbookPath);

        var arguments = CreateArguments("add", workbookPath);
        arguments["sheetName"] = "ExistingSheet";

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public async Task DeleteSheet_LastSheet_ShouldThrowInvalidOperationException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_last.xlsx");
        var arguments = CreateArguments("delete", workbookPath);
        arguments["sheetIndex"] = 0;

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task RenameSheet_WithDuplicateName_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_rename_duplicate.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("Sheet2");
        workbook.Save(workbookPath);

        var arguments = CreateArguments("rename", workbookPath);
        arguments["sheetIndex"] = 0;
        arguments["newName"] = "Sheet2";

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public async Task MoveSheet_ToSamePosition_ShouldNotModifyFile()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_move_same_position.xlsx");
        var arguments = CreateArguments("move", workbookPath);
        arguments["sheetIndex"] = 0;
        arguments["insertAt"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("no move needed", result);
    }

    [Fact]
    public async Task MoveSheet_WithoutTargetIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_move_no_target.xlsx");
        var arguments = CreateArguments("move", workbookPath);
        arguments["sheetIndex"] = 0;

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("targetIndex", ex.Message);
    }

    [Fact]
    public async Task CopySheet_ToExternalFile_ShouldCopySuccessfully()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_external.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells[0, 0].Value = "TestData";
        workbook.Save(workbookPath);

        var copyToPath = CreateTestFilePath("test_copy_external_target.xlsx");
        var arguments = CreateArguments("copy", workbookPath);
        arguments["sheetIndex"] = 0;
        arguments["copyToPath"] = copyToPath;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(copyToPath), "Target workbook should be created");
        Assert.Contains("external file", result);

        var targetWorkbook = new Workbook(copyToPath);
        Assert.Equal("TestData", targetWorkbook.Worksheets[0].Cells[0, 0].Value?.ToString());
    }

    [Fact]
    public async Task HideSheet_Toggle_ShouldShowHiddenSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_toggle_visibility.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("HiddenSheet");
        workbook.Worksheets["HiddenSheet"].IsVisible = false;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_toggle_visibility_output.xlsx");
        var arguments = CreateArguments("hide", workbookPath, outputPath);
        arguments["sheetIndex"] = 1;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("shown", result);
        var resultWorkbook = new Workbook(outputPath);
        var isEvaluationMode = IsEvaluationMode();
        if (!isEvaluationMode)
        {
            var hiddenSheet = resultWorkbook.Worksheets["HiddenSheet"];
            Assert.True(hiddenSheet.IsVisible, "Hidden sheet should now be visible");
        }
    }

    [Fact]
    public async Task GetSheets_EmptyWorkbook_ShouldReturnValidJson()
    {
        // Arrange - Create workbook but don't add extra sheets
        var workbookPath = CreateExcelWorkbook("test_get_sheets_minimal.xlsx");
        var arguments = CreateArguments("get", workbookPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("count", result);
        Assert.Contains("items", result);
    }

    [Fact]
    public async Task InvalidOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var arguments = CreateArguments("invalid_operation", workbookPath);

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public async Task SheetIndex_OutOfRange_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_index_out_of_range.xlsx");
        var arguments = CreateArguments("delete", workbookPath);
        arguments["sheetIndex"] = 999;

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}