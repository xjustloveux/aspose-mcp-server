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

    private string CreateWorkbookWithNamedRange(string fileName, string rangeName, string rangeAddress)
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        for (var row = 0; row < 5; row++)
        for (var col = 0; col < 5; col++)
            worksheet.Cells[row, col].Value = $"R{row}C{col}";
        var parts = rangeAddress.Split(':');
        var range = parts.Length > 1
            ? worksheet.Cells.CreateRange(parts[0], parts[1])
            : worksheet.Cells.CreateRange(parts[0], parts[0]);
        range.Name = rangeName;
        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddNamedRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, name: "TestRange", range: "A1:C5", outputPath: outputPath);
        Assert.StartsWith("Named range 'TestRange' added", result);
        Assert.Contains("reference:", result);
        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets.Names["TestRange"]);
    }

    [Fact]
    public void Add_WithComment_ShouldAddComment()
    {
        var workbookPath = CreateExcelWorkbook("test_add_comment.xlsx");
        var outputPath = CreateTestFilePath("test_add_comment_output.xlsx");
        var result = _tool.Execute("add", workbookPath, name: "CommentedRange", range: "A1:B2",
            comment: "This is a test range", outputPath: outputPath);
        Assert.StartsWith("Named range 'CommentedRange' added", result);
        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["CommentedRange"];
        Assert.NotNull(namedRange);
        Assert.Equal("This is a test range", namedRange.Comment);
    }

    [Fact]
    public void Add_SingleCell_ShouldAddRange()
    {
        var workbookPath = CreateExcelWorkbook("test_add_single.xlsx");
        var outputPath = CreateTestFilePath("test_add_single_output.xlsx");
        var result = _tool.Execute("add", workbookPath, name: "SingleCell", range: "A1", outputPath: outputPath);
        Assert.StartsWith("Named range 'SingleCell' added", result);
        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets.Names["SingleCell"]);
    }

    [Fact]
    public void Add_WithSheetReference_ShouldAddToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add_sheetref.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("DataSheet");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_sheetref_output.xlsx");
        var result = _tool.Execute("add", workbookPath, name: "SheetRange",
            range: "DataSheet!A1:C5", outputPath: outputPath);
        Assert.StartsWith("Named range 'SheetRange' added", result);
        Assert.Contains("DataSheet", result);
        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["SheetRange"];
        Assert.NotNull(namedRange);
        Assert.Contains("DataSheet", namedRange.RefersTo);
    }

    [SkippableFact]
    public void Add_WithSheetIndex_ShouldAddToCorrectSheet()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Adding sheet exceeds evaluation limit");
        var workbookPath = CreateExcelWorkbook("test_add_sheetindex.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_sheetindex_output.xlsx");
        var result = _tool.Execute("add", workbookPath, name: "Sheet2Range",
            range: "A1:C5", sheetIndex: 1, outputPath: outputPath);
        Assert.StartsWith("Named range 'Sheet2Range' added", result);
        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["Sheet2Range"];
        Assert.NotNull(namedRange);
        Assert.Contains("Sheet2", namedRange.RefersTo);
    }

    [Fact]
    public void Delete_ShouldDeleteNamedRange()
    {
        var workbookPath = CreateWorkbookWithNamedRange("test_delete.xlsx", "RangeToDelete", "A1:B2");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, name: "RangeToDelete", outputPath: outputPath);
        Assert.StartsWith("Named range 'RangeToDelete' deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Null(workbook.Worksheets.Names["RangeToDelete"]);
    }

    [Fact]
    public void Get_ShouldReturnAllNamedRanges()
    {
        var workbookPath = CreateExcelWorkbook("test_get.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells.CreateRange("A1", "B2").Name = "Range1";
            wb.Worksheets[0].Cells.CreateRange("C1", "D2").Name = "Range2";
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("Range1", result);
        Assert.Contains("Range2", result);
    }

    [Fact]
    public void Get_NoNamedRanges_ShouldReturnEmptyMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal("No named ranges found", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public void Get_ShouldIncludeAllProperties()
    {
        var workbookPath = CreateExcelWorkbook("test_get_props.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var range = wb.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "DetailedRange";
            wb.Worksheets.Names["DetailedRange"].Comment = "Test comment";
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        var firstItem = json.RootElement.GetProperty("items")[0];
        Assert.True(firstItem.TryGetProperty("name", out _));
        Assert.True(firstItem.TryGetProperty("reference", out _));
        Assert.True(firstItem.TryGetProperty("comment", out _));
        Assert.True(firstItem.TryGetProperty("isVisible", out _));
        Assert.Equal("DetailedRange", firstItem.GetProperty("name").GetString());
        Assert.Equal("Test comment", firstItem.GetProperty("comment").GetString());
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, name: $"Range_{operation}",
            range: "A1:B2", outputPath: outputPath);
        Assert.Contains("added", result);
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

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var workbookPath = CreateWorkbookWithNamedRange($"test_case_del_{operation}.xlsx", "TestRange", "A1:B2");
        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, name: "TestRange", outputPath: outputPath);
        Assert.Contains("deleted", result);
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
    public void Add_WithMissingName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_name.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:B2"));
        Assert.Contains("name", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, name: "TestRange"));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithDuplicateName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithNamedRange("test_add_duplicate.xlsx", "ExistingRange", "A1:B2");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, name: "ExistingRange", range: "C1:D2"));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, name: "TestRange", range: "A1:B2", sheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidSheetReference_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheetref.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, name: "TestRange", range: "NonExistentSheet!A1:B2"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_missing_name.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath));
        Assert.Contains("name", ex.Message.ToLower());
    }

    [Fact]
    public void Delete_WithNonExistentName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_nonexistent.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, name: "NonExistentRange"));
        Assert.Contains("does not exist", ex.Message);
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
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_add.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, name: "InMemoryRange", range: "A1:C3");
        Assert.StartsWith("Named range 'InMemoryRange' added", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotNull(workbook.Worksheets.Names["InMemoryRange"]);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithNamedRange("test_session_delete.xlsx", "RangeToDelete", "A1:B2");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, name: "RangeToDelete");
        Assert.StartsWith("Named range 'RangeToDelete' deleted", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Null(workbook.Worksheets.Names["RangeToDelete"]);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithNamedRange("test_session_get.xlsx", "SessionRange", "A1:B2");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("SessionRange", result);
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
        var workbookPath2 = CreateWorkbookWithNamedRange("test_session_file.xlsx", "SessionRange", "A1:B2");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        Assert.Contains("SessionRange", result);
    }

    #endregion
}