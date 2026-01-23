using AsposeMcpServer.Handlers.Excel.Comment;
using AsposeMcpServer.Results.Excel.Comment;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.True(result.Items.Count > 0);
        var firstItem = result.Items[0];
        Assert.NotNull(firstItem.Cell);
        Assert.NotNull(firstItem.Note);
        Assert.NotNull(firstItem.Author);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Items);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.Equal(2, result.Count);
    }

    [Fact]
    public void Execute_WithNoComments_ReturnsEmptyList()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.Equal(0, result.Count);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.Equal(1, result.Count);
        Assert.Equal("A1", result.Cell);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.Equal(0, result.Count);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.Equal(1, result.SheetIndex);
        Assert.Equal(1, result.Count);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("A1");
        workbook.Worksheets[0].Comments[idx].Note = "Comment on first sheet";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsExcelResult>(res);

        Assert.Equal(0, result.SheetIndex);
    }

    #endregion
}
