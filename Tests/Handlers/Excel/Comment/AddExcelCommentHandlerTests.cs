using AsposeMcpServer.Handlers.Excel.Comment;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Comment;

public class AddExcelCommentHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsCellAndSheetInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B5" },
            { "comment", "Test" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B5", result);
        Assert.Contains("sheet", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsComment()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Test comment" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(workbook.Worksheets[0].Comments["A1"]);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1", "Comment on A1")]
    [InlineData("B2", "Comment on B2")]
    [InlineData("C10", "Comment on C10")]
    public void Execute_AddsCommentToVariousCells(string cell, string comment)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", cell },
            { "comment", comment }
        });

        _handler.Execute(context, parameters);

        var commentObj = workbook.Worksheets[0].Comments[cell];
        Assert.NotNull(commentObj);
        Assert.Equal(comment, commentObj.Note);
    }

    #endregion

    #region Author

    [Fact]
    public void Execute_WithAuthor_SetsAuthor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Test comment" },
            { "author", "John Doe" }
        });

        _handler.Execute(context, parameters);

        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.Equal("John Doe", comment.Author);
    }

    [Fact]
    public void Execute_WithoutAuthor_UsesDefaultAuthor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Test comment" }
        });

        _handler.Execute(context, parameters);

        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.NotNull(comment.Author);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_AddsToSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "cell", "A1" },
            { "comment", "Comment on Sheet2" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 1", result);
        Assert.NotNull(workbook.Worksheets[1].Comments["A1"]);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_AddsToFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Comment on first sheet" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 0", result);
        Assert.NotNull(workbook.Worksheets[0].Comments["A1"]);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "comment", "Test comment" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutComment_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("comment", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "" },
            { "comment", "Test comment" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidCellAddress_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "InvalidCell" },
            { "comment", "Test comment" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
