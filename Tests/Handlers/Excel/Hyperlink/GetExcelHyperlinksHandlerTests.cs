using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Hyperlink;
using AsposeMcpServer.Results.Excel.Hyperlink;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Empty(result.Items);
    }

    [Fact]
    public void Execute_WithNoHyperlinks_ReturnsMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal("No hyperlinks found", result.Message);
    }

    [Fact]
    public void Execute_WithHyperlinks_ReturnsCount()
    {
        var workbook = CreateWorkbookWithHyperlinks(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_WithHyperlinks_ReturnsItems()
    {
        var workbook = CreateWorkbookWithHyperlinks(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.NotNull(result.Items);
        Assert.Equal(2, result.Items.Count);
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.NotNull(result.WorksheetName);
    }

    #endregion

    #region Hyperlink Details

    [Fact]
    public void Execute_ReturnsHyperlinkIndex()
    {
        var workbook = CreateWorkbookWithHyperlinks(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal(0, result.Items[0].Index);
        Assert.Equal(1, result.Items[1].Index);
    }

    [Fact]
    public void Execute_ReturnsHyperlinkCell()
    {
        var workbook = CreateWorkbookWithHyperlink("B2", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal("B2", result.Items[0].Cell);
    }

    [Fact]
    public void Execute_ReturnsHyperlinkUrl()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://test.example.com");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal("https://test.example.com", result.Items[0].Url);
    }

    [Fact]
    public void Execute_ReturnsDisplayText()
    {
        var workbook = CreateWorkbookWithHyperlinkAndDisplayText("A1", "https://example.com", "Click Here");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal("Click Here", result.Items[0].DisplayText);
    }

    [Fact]
    public void Execute_ReturnsAreaInfo()
    {
        var workbook = CreateWorkbookWithHyperlinks(1);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.NotNull(result.Items[0].Area);
        Assert.NotNull(result.Items[0].Area.StartCell);
        Assert.NotNull(result.Items[0].Area.EndCell);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal(1, result.Count);
        Assert.Equal("Sheet2", result.WorksheetName);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateWorkbookWithHyperlinks(2);
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksExcelResult>(res);

        Assert.Equal(2, result.Count);
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
