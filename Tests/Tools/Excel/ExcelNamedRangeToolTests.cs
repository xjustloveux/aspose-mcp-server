using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelNamedRangeToolTests : ExcelTestBase
{
    private readonly ExcelNamedRangeTool _tool;

    public ExcelNamedRangeToolTests()
    {
        _tool = new ExcelNamedRangeTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddNamedRange_ShouldAddNamedRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_named_range.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_named_range_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            name: "TestRange",
            range: "A1:C5",
            outputPath: outputPath);
        Assert.Contains("Named range 'TestRange' added", result);
        Assert.Contains("reference:", result);

        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets.Names["TestRange"]);
    }

    [Fact]
    public void AddNamedRange_WithComment_ShouldAddComment()
    {
        var workbookPath = CreateExcelWorkbook("test_add_named_range_comment.xlsx");
        var outputPath = CreateTestFilePath("test_add_named_range_comment_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            name: "CommentedRange",
            range: "A1:B2",
            comment: "This is a test range",
            outputPath: outputPath);
        Assert.Contains("Named range 'CommentedRange' added", result);

        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["CommentedRange"];
        Assert.NotNull(namedRange);
        Assert.Equal("This is a test range", namedRange.Comment);
    }

    [Fact]
    public void AddNamedRange_SingleCell_ShouldAddRange()
    {
        var workbookPath = CreateExcelWorkbook("test_add_single_cell.xlsx");
        var outputPath = CreateTestFilePath("test_add_single_cell_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            name: "SingleCell",
            range: "A1",
            outputPath: outputPath);
        Assert.Contains("Named range 'SingleCell' added", result);

        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets.Names["SingleCell"]);
    }

    [Fact]
    public void AddNamedRange_WithSheetReference_ShouldAddToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add_sheet_ref.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("DataSheet");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_sheet_ref_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            name: "SheetRange",
            range: "DataSheet!A1:C5",
            outputPath: outputPath);
        Assert.Contains("Named range 'SheetRange' added", result);
        Assert.Contains("DataSheet", result);

        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["SheetRange"];
        Assert.NotNull(namedRange);
        Assert.Contains("DataSheet", namedRange.RefersTo);
    }

    [Fact]
    public void AddNamedRange_DuplicateName_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_duplicate.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var range = wb.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "ExistingRange";
            wb.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            name: "ExistingRange",
            range: "C1:D2"));
        Assert.Contains("already exists", exception.Message);
    }

    [Fact]
    public void AddNamedRange_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            name: "InvalidSheet",
            range: "A1:B2",
            sheetIndex: 99));
    }

    [Fact]
    public void AddNamedRange_InvalidSheetReference_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet_ref.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            name: "InvalidRef",
            range: "NonExistentSheet!A1:B2"));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void DeleteNamedRange_ShouldDeleteNamedRange()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_named_range.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var range = workbook.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "RangeToDelete";
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_named_range_output.xlsx");
        var result = _tool.Execute(
            "delete",
            workbookPath,
            name: "RangeToDelete",
            outputPath: outputPath);
        Assert.Contains("Named range 'RangeToDelete' deleted", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Null(resultWorkbook.Worksheets.Names["RangeToDelete"]);
    }

    [Fact]
    public void DeleteNamedRange_NonExistent_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_nonexistent.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            workbookPath,
            name: "NonExistentRange"));
        Assert.Contains("does not exist", exception.Message);
    }

    [Fact]
    public void GetNamedRanges_ShouldReturnAllNamedRanges()
    {
        var workbookPath = CreateExcelWorkbook("test_get_named_ranges.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var range1 = workbook.Worksheets[0].Cells.CreateRange("A1", "B2");
            range1.Name = "Range1";
            var range2 = workbook.Worksheets[0].Cells.CreateRange("C1", "D2");
            range2.Name = "Range2";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(2, root.GetProperty("count").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(2, items.GetArrayLength());

        Assert.Contains("Range1", result);
        Assert.Contains("Range2", result);
    }

    [Fact]
    public void GetNamedRanges_WithNoNamedRanges_ShouldReturnEmptyMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty_named_ranges.xlsx");
        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No named ranges found", root.GetProperty("message").GetString());
    }

    [Fact]
    public void GetNamedRanges_ShouldIncludeAllProperties()
    {
        var workbookPath = CreateExcelWorkbook("test_get_properties.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var range = workbook.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "DetailedRange";
            var namedRange = workbook.Worksheets.Names["DetailedRange"];
            namedRange.Comment = "Test comment";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var firstItem = items[0];

        Assert.True(firstItem.TryGetProperty("name", out _));
        Assert.True(firstItem.TryGetProperty("reference", out _));
        Assert.True(firstItem.TryGetProperty("comment", out _));
        Assert.True(firstItem.TryGetProperty("isVisible", out _));
        Assert.Equal("DetailedRange", firstItem.GetProperty("name").GetString());
        Assert.Equal("Test comment", firstItem.GetProperty("comment").GetString());
    }

    [SkippableFact]
    public void AddNamedRange_WithSheetIndex_ShouldAddToCorrectSheet()
    {
        // Skip in evaluation mode - adding sheet exceeds evaluation limit
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Adding sheet exceeds evaluation limit");
        var workbookPath = CreateExcelWorkbook("test_add_with_sheet_index.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_with_sheet_index_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            name: "Sheet2Range",
            range: "A1:C5",
            sheetIndex: 1,
            outputPath: outputPath);
        Assert.Contains("Named range 'Sheet2Range' added", result);

        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["Sheet2Range"];
        Assert.NotNull(namedRange);
        Assert.Contains("Sheet2", namedRange.RefersTo);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid",
            workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void AddNamedRange_MissingName_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_name.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            range: "A1:B2"));
        Assert.Contains("name", exception.Message.ToLower());
    }

    [Fact]
    public void AddNamedRange_MissingRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_range.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            name: "TestRange"));
        Assert.Contains("range", exception.Message.ToLower());
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetNamedRanges_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_ranges.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var range = workbook.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "SessionRange";
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "get",
            sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
        Assert.Contains("SessionRange", result);
    }

    [Fact]
    public void AddNamedRange_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_add_range.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            name: "InMemoryRange",
            range: "A1:C3");
        Assert.Contains("Named range 'InMemoryRange' added", result);

        // Verify in-memory workbook has the named range
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotNull(workbook.Worksheets.Names["InMemoryRange"]);
    }

    [Fact]
    public void DeleteNamedRange_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_delete_range.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var range = wb.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "RangeToDelete";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "delete",
            sessionId: sessionId,
            name: "RangeToDelete");
        Assert.Contains("Named range 'RangeToDelete' deleted", result);

        // Verify in-memory workbook has no named range
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Null(workbook.Worksheets.Names["RangeToDelete"]);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}