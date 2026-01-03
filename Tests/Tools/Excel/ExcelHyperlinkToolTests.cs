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

    #region General Tests

    [Fact]
    public void AddHyperlink_ShouldAddHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_add_hyperlink.xlsx");
        var outputPath = CreateTestFilePath("test_add_hyperlink_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            cell: "A1",
            url: "https://example.com",
            displayText: "Click here",
            outputPath: outputPath);
        Assert.Contains("Hyperlink added to A1", result);
        Assert.Contains("https://example.com", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Hyperlinks);
        Assert.Equal("https://example.com", worksheet.Hyperlinks[0].Address);
        Assert.Equal("Click here", worksheet.Hyperlinks[0].TextToDisplay);
    }

    [Fact]
    public void AddHyperlink_WithoutDisplayText_ShouldAddHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_add_no_display.xlsx");
        var outputPath = CreateTestFilePath("test_add_no_display_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            cell: "B2",
            url: "https://test.com",
            outputPath: outputPath);
        Assert.Contains("Hyperlink added to B2", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Hyperlinks);
    }

    [Fact]
    public void GetHyperlinks_ShouldReturnAllHyperlinks()
    {
        var workbookPath = CreateExcelWorkbook("test_get_hyperlinks.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Hyperlinks.Add("A1", 1, 1, "https://test1.com");
            worksheet.Hyperlinks.Add("B2", 1, 1, "https://test2.com");
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

        var first = items[0];
        Assert.Equal("A1", first.GetProperty("cell").GetString());
        Assert.Equal("https://test1.com", first.GetProperty("url").GetString());
    }

    [Fact]
    public void GetHyperlinks_EmptyWorksheet_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No hyperlinks found", root.GetProperty("message").GetString());
    }

    [Fact]
    public void EditHyperlink_ByCell_ShouldModifyHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_hyperlink.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://old.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            cell: "A1",
            url: "https://new.com",
            displayText: "New Link",
            outputPath: outputPath);
        Assert.Contains("edited", result);
        Assert.Contains("url=https://new.com", result);

        using var resultWorkbook = new Workbook(outputPath);
        var hyperlink = resultWorkbook.Worksheets[0].Hyperlinks[0];
        Assert.Equal("https://new.com", hyperlink.Address);
        Assert.Equal("New Link", hyperlink.TextToDisplay);
    }

    [Fact]
    public void EditHyperlink_ByIndex_ShouldModifyHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_by_index.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://first.com");
            workbook.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://second.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_by_index_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            hyperlinkIndex: 1,
            url: "https://modified.com",
            outputPath: outputPath);
        Assert.Contains("edited", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("https://first.com", resultWorkbook.Worksheets[0].Hyperlinks[0].Address);
        Assert.Equal("https://modified.com", resultWorkbook.Worksheets[0].Hyperlinks[1].Address);
    }

    [Fact]
    public void DeleteHyperlink_ByCell_ShouldDeleteHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_hyperlink.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://delete.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.xlsx");
        var result = _tool.Execute(
            "delete",
            workbookPath,
            cell: "A1",
            outputPath: outputPath);
        Assert.Contains("deleted", result);
        Assert.Contains("0 hyperlinks remaining", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void DeleteHyperlink_ByIndex_ShouldDeleteHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_by_index.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://first.com");
            workbook.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://second.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_by_index_output.xlsx");
        var result = _tool.Execute(
            "delete",
            workbookPath,
            hyperlinkIndex: 0,
            outputPath: outputPath);
        Assert.Contains("deleted", result);
        Assert.Contains("1 hyperlinks remaining", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Single(resultWorkbook.Worksheets[0].Hyperlinks);
        Assert.Equal("https://second.com", resultWorkbook.Worksheets[0].Hyperlinks[0].Address);
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
    public void AddHyperlink_CellAlreadyHasHyperlink_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_existing.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://existing.com");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            cell: "A1",
            url: "https://new.com"));
        Assert.Contains("already has a hyperlink", exception.Message);
    }

    [Fact]
    public void EditHyperlink_NoCellOrIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_missing.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            workbookPath,
            url: "https://new.com"));
        Assert.Contains("Either 'hyperlinkIndex' or 'cell' is required", exception.Message);
    }

    [Fact]
    public void EditHyperlink_CellNotFound_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_not_found.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            workbookPath,
            cell: "Z99",
            url: "https://new.com"));
        Assert.Contains("No hyperlink found at cell", exception.Message);
    }

    [Fact]
    public void DeleteHyperlink_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://test.com");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            workbookPath,
            hyperlinkIndex: 99));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void AddHyperlink_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            sheetIndex: 99,
            cell: "A1",
            url: "https://test.com"));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Hyperlinks.Add("A1", 1, 1, "https://test1.com");
            worksheet.Hyperlinks.Add("B2", 1, 1, "https://test2.com");
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "get",
            sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            cell: "A1",
            url: "https://session-test.com",
            displayText: "Session Link");
        Assert.Contains("Hyperlink added to A1", result);

        // Verify in-memory workbook has the hyperlink
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].Hyperlinks);
        Assert.Equal("https://session-test.com", workbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_edit.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://old.com");
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "edit",
            sessionId: sessionId,
            cell: "A1",
            url: "https://new-session.com");
        Assert.Contains("edited", result);

        // Verify in-memory workbook has the updated hyperlink
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("https://new-session.com", sessionWorkbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_delete.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://delete.com");
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "delete",
            sessionId: sessionId,
            cell: "A1");
        Assert.Contains("deleted", result);

        // Verify in-memory workbook has no hyperlinks
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(sessionWorkbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Add_WithSessionId_ShouldNotModifyOriginalFile()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add_original.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute(
            "add",
            sessionId: sessionId,
            cell: "A1",
            url: "https://session-only.com");

        // Assert - original file should not have the hyperlink
        using var originalWorkbook = new Workbook(workbookPath);
        Assert.Empty(originalWorkbook.Worksheets[0].Hyperlinks);

        // But session workbook should have the hyperlink
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(sessionWorkbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}