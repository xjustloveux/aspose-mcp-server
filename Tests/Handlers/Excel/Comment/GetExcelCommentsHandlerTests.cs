using System.Text.Json;
using AsposeMcpServer.Handlers.Excel.Comment;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Comment;

public class GetExcelCommentsHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelCommentsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Comment Properties

    [Fact]
    public void Execute_ReturnsCommentProperties()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add("A1");
        sheet.Comments[idx].Note = "Test note";
        sheet.Comments[idx].Author = "Test Author";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items.GetArrayLength() > 0);
        var firstItem = items[0];
        Assert.True(firstItem.TryGetProperty("cell", out _));
        Assert.True(firstItem.TryGetProperty("note", out _));
        Assert.True(firstItem.TryGetProperty("author", out _));
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyWorkbook()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("A1");
        workbook.Worksheets[0].Comments[idx].Note = "Test";
        var initialCount = workbook.Worksheets[0].Comments.Count;
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, workbook.Worksheets[0].Comments.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsCommentsInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var commentIndex = sheet.Comments.Add("A1");
        sheet.Comments[commentIndex].Note = "Test comment";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("items", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx1 = sheet.Comments.Add("A1");
        sheet.Comments[idx1].Note = "Comment 1";
        var idx2 = sheet.Comments.Add("B2");
        sheet.Comments[idx2].Note = "Comment 2";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithNoComments_ReturnsEmptyList()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region Get Specific Cell Comment

    [Fact]
    public void Execute_WithCell_ReturnsSpecificComment()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add("A1");
        sheet.Comments[idx].Note = "Specific comment";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal("A1", json.RootElement.GetProperty("cell").GetString());
    }

    [Fact]
    public void Execute_WithCellNoComment_ReturnsEmptyResult()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_GetsFromSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var idx = workbook.Worksheets[1].Comments.Add("A1");
        workbook.Worksheets[1].Comments[idx].Note = "Comment on Sheet2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("sheetIndex").GetInt32());
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("A1");
        workbook.Worksheets[0].Comments[idx].Note = "Comment on first sheet";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("sheetIndex").GetInt32());
    }

    #endregion
}
