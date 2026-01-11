using AsposeMcpServer.Handlers.Excel.Comment;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Comment;

public class EditExcelCommentHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsCellAndSheetInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("B5");
        workbook.Worksheets[0].Comments[idx].Note = "Original";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B5" },
            { "comment", "Updated" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B5", result);
        Assert.Contains("sheet", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsComment()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add("A1");
        sheet.Comments[idx].Note = "Original comment";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Updated comment" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Updated comment", sheet.Comments["A1"].Note);
        AssertModified(context);
    }

    [Theory]
    [InlineData("New content 1")]
    [InlineData("New content 2")]
    [InlineData("Special chars: !@#$%")]
    public void Execute_UpdatesCommentContent(string newContent)
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add("A1");
        sheet.Comments[idx].Note = "Original";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", newContent }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(newContent, sheet.Comments["A1"].Note);
    }

    #endregion

    #region Author

    [Fact]
    public void Execute_WithAuthor_UpdatesAuthor()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add("A1");
        sheet.Comments[idx].Note = "Comment";
        sheet.Comments[idx].Author = "Original Author";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Updated comment" },
            { "author", "New Author" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("New Author", sheet.Comments["A1"].Author);
    }

    [Fact]
    public void Execute_WithoutAuthor_PreservesOriginalAuthor()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add("A1");
        sheet.Comments[idx].Note = "Comment";
        sheet.Comments[idx].Author = "Original Author";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Updated comment" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Original Author", sheet.Comments["A1"].Author);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_EditsOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var idx = workbook.Worksheets[1].Comments.Add("A1");
        workbook.Worksheets[1].Comments[idx].Note = "Original";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "cell", "A1" },
            { "comment", "Updated on Sheet2" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 1", result);
        Assert.Equal("Updated on Sheet2", workbook.Worksheets[1].Comments["A1"].Note);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_EditsOnFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("A1");
        workbook.Worksheets[0].Comments[idx].Note = "Original";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "Updated on first sheet" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 0", result);
        Assert.Equal("Updated on first sheet", workbook.Worksheets[0].Comments["A1"].Note);
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
            { "comment", "Updated comment" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutComment_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("A1");
        workbook.Worksheets[0].Comments[idx].Note = "Original";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("comment", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_OnCellWithoutComment_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "comment", "New comment" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("No comment found", ex.Message);
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
