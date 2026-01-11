using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Hyperlink;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Hyperlink;

public class GetExcelHyperlinksHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelHyperlinksHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Get All Hyperlinks

    [Fact]
    public void Execute_WithNoHyperlinks_ReturnsEmptyResult()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.Contains("\"count\": 0", result);
    }

    [Fact]
    public void Execute_WithNoHyperlinks_ReturnsMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No hyperlinks found", result);
    }

    [Fact]
    public void Execute_WithHyperlinks_ReturnsCount()
    {
        var workbook = CreateWorkbookWithHyperlinks(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 3", result);
    }

    [Fact]
    public void Execute_WithHyperlinks_ReturnsJsonFormat()
    {
        var workbook = CreateWorkbookWithHyperlinks(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("worksheetName", result);
    }

    #endregion

    #region Hyperlink Details

    [Fact]
    public void Execute_ReturnsHyperlinkIndex()
    {
        var workbook = CreateWorkbookWithHyperlinks(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"index\": 0", result);
        Assert.Contains("\"index\": 1", result);
    }

    [Fact]
    public void Execute_ReturnsHyperlinkCell()
    {
        var workbook = CreateWorkbookWithHyperlink("B2", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"cell\": \"B2\"", result);
    }

    [Fact]
    public void Execute_ReturnsHyperlinkUrl()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://test.example.com");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("https://test.example.com", result);
    }

    [Fact]
    public void Execute_ReturnsDisplayText()
    {
        var workbook = CreateWorkbookWithHyperlinkAndDisplayText("A1", "https://example.com", "Click Here");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Click Here", result);
    }

    [Fact]
    public void Execute_ReturnsAreaInfo()
    {
        var workbook = CreateWorkbookWithHyperlinks(1);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("area", result);
        Assert.Contains("startCell", result);
        Assert.Contains("endCell", result);
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_GetsFromCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Hyperlinks.Add("A1", 1, 1, "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 1", result);
        Assert.Contains("Sheet2", result);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateWorkbookWithHyperlinks(2);
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 2", result);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithHyperlink(string cell, string url)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Hyperlinks.Add(cell, 1, 1, url);
        return workbook;
    }

    private static Workbook CreateWorkbookWithHyperlinkAndDisplayText(string cell, string url, string displayText)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Hyperlinks.Add(cell, 1, 1, url);
        sheet.Hyperlinks[idx].TextToDisplay = displayText;
        return workbook;
    }

    private static Workbook CreateWorkbookWithHyperlinks(int count)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        for (var i = 0; i < count; i++) sheet.Hyperlinks.Add($"A{i + 1}", 1, 1, $"https://example{i}.com");
        return workbook;
    }

    #endregion
}
