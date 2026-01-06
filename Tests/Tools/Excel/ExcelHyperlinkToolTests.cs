using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelHyperlinkToolTests : ExcelTestBase
{
    private readonly ExcelHyperlinkTool _tool;

    public ExcelHyperlinkToolTests()
    {
        _tool = new ExcelHyperlinkTool(SessionManager);
    }

    private string CreateWorkbookWithHyperlink(string fileName, string cell = "A1", string url = "https://test.com")
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        workbook.Worksheets[0].Hyperlinks.Add(cell, 1, 1, url);
        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, cell: "A1",
            url: "https://example.com", displayText: "Click here", outputPath: outputPath);
        Assert.Contains("Hyperlink added to A1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Hyperlinks);
        Assert.Equal("https://example.com", workbook.Worksheets[0].Hyperlinks[0].Address);
        Assert.Equal("Click here", workbook.Worksheets[0].Hyperlinks[0].TextToDisplay);
    }

    [Fact]
    public void Add_WithoutDisplayText_ShouldAddHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_add_no_text.xlsx");
        var outputPath = CreateTestFilePath("test_add_no_text_output.xlsx");
        var result = _tool.Execute("add", workbookPath, cell: "B2",
            url: "https://test.com", outputPath: outputPath);
        Assert.Contains("Hyperlink added to B2", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Add_WithSheetIndex_ShouldAddToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_sheet_output.xlsx");
        var result = _tool.Execute("add", workbookPath, sheetIndex: 1,
            cell: "A1", url: "https://sheet2.com", outputPath: outputPath);
        Assert.Contains("Hyperlink added to A1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Hyperlinks);
        Assert.Single(workbook.Worksheets[1].Hyperlinks);
    }

    [Fact]
    public void Edit_ByCell_ShouldModifyHyperlink()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_edit_cell.xlsx", "A1", "https://old.com");
        var outputPath = CreateTestFilePath("test_edit_cell_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, cell: "A1",
            url: "https://new.com", displayText: "New Link", outputPath: outputPath);
        Assert.Contains("Hyperlink at", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("https://new.com", workbook.Worksheets[0].Hyperlinks[0].Address);
        Assert.Equal("New Link", workbook.Worksheets[0].Hyperlinks[0].TextToDisplay);
    }

    [Fact]
    public void Edit_ByIndex_ShouldModifyHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_index.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://first.com");
            wb.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://second.com");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_index_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, hyperlinkIndex: 1,
            url: "https://modified.com", outputPath: outputPath);
        Assert.Contains("Hyperlink at", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("https://first.com", workbook.Worksheets[0].Hyperlinks[0].Address);
        Assert.Equal("https://modified.com", workbook.Worksheets[0].Hyperlinks[1].Address);
    }

    [Fact]
    public void Edit_DisplayTextOnly_ShouldModifyDisplayText()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_edit_text.xlsx", "A1", "https://keep.com");
        var outputPath = CreateTestFilePath("test_edit_text_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, cell: "A1",
            displayText: "New Display", outputPath: outputPath);
        Assert.Contains("Hyperlink at", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("https://keep.com", workbook.Worksheets[0].Hyperlinks[0].Address);
        Assert.Equal("New Display", workbook.Worksheets[0].Hyperlinks[0].TextToDisplay);
    }

    [Fact]
    public void Delete_ByCell_ShouldDeleteHyperlink()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_delete_cell.xlsx");
        var outputPath = CreateTestFilePath("test_delete_cell_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, cell: "A1", outputPath: outputPath);
        Assert.Contains("deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Delete_ByIndex_ShouldDeleteHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_index.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://first.com");
            wb.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://second.com");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_index_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, hyperlinkIndex: 0, outputPath: outputPath);
        Assert.Contains("deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Hyperlinks);
        Assert.Equal("https://second.com", workbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Get_ShouldReturnAllHyperlinks()
    {
        var workbookPath = CreateExcelWorkbook("test_get.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://test1.com");
            wb.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://test2.com");
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(2, root.GetProperty("count").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(2, items.GetArrayLength());
        Assert.Equal("A1", items[0].GetProperty("cell").GetString());
        Assert.Equal("https://test1.com", items[0].GetProperty("url").GetString());
    }

    [Fact]
    public void Get_Empty_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal("No hyperlinks found", json.RootElement.GetProperty("message").GetString());
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1",
            url: "https://test.com", outputPath: outputPath);
        Assert.Contains("Hyperlink added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("\"count\":", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var workbookPath = CreateWorkbookWithHyperlink($"test_case_del_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1", outputPath: outputPath);
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
    public void Add_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, url: "https://test.com"));
        Assert.Contains("cell", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithMissingUrl_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_url.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "A1"));
        Assert.Contains("url", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithExistingHyperlink_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_add_existing.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "A1", url: "https://new.com"));
        Assert.Contains("already has a hyperlink", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetIndex: 99, cell: "A1", url: "https://test.com"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingCellAndIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_missing.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, url: "https://new.com"));
        Assert.Contains("Either 'hyperlinkIndex' or 'cell' is required", ex.Message);
    }

    [Fact]
    public void Edit_WithCellNotFound_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_notfound.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, cell: "Z99", url: "https://new.com"));
        Assert.Contains("No hyperlink found at cell", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_edit_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, hyperlinkIndex: 99, url: "https://new.com"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingCellAndIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_missing.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", workbookPath));
        Assert.Contains("Either 'hyperlinkIndex' or 'cell' is required", ex.Message);
    }

    [Fact]
    public void Delete_WithCellNotFound_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_notfound.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, cell: "Z99"));
        Assert.Contains("No hyperlink found at cell", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_delete_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, hyperlinkIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_invalid_sheet.xlsx");
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
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, cell: "A1",
            url: "https://session.com", displayText: "Session Link");
        Assert.Contains("Hyperlink added to A1", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].Hyperlinks);
        Assert.Equal("https://session.com", workbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_session_edit.xlsx", "A1", "https://old.com");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("edit", sessionId: sessionId, cell: "A1", url: "https://new-session.com");
        Assert.Contains("Hyperlink at", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("https://new-session.com", workbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, cell: "A1");
        Assert.Contains("deleted", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://test1.com");
            wb.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://test2.com");
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
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
        var workbookPath2 = CreateWorkbookWithHyperlink("test_session_file.xlsx", "A1", "https://session.com");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        Assert.Contains("https://session.com", result);
    }

    #endregion
}