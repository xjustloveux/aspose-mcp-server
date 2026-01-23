using AsposeMcpServer.Handlers.Excel.Comment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Comment;

public class DeleteExcelCommentHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Preserves Other Comments

    [Fact]
    public void Execute_PreservesOtherComments()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx1 = sheet.Comments.Add("A1");
        sheet.Comments[idx1].Note = "Comment 1";
        var idx2 = sheet.Comments.Add("B2");
        sheet.Comments[idx2].Note = "Comment 2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Null(sheet.Comments["A1"]);
        Assert.NotNull(sheet.Comments["B2"]);
        Assert.Equal("Comment 2", sheet.Comments["B2"].Note);
    }

    #endregion

    #region Delete Non-Existent Comment

    [Fact]
    public void Execute_OnCellWithoutComment_SucceedsWithoutError()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsCellAndSheetInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("B5");
        workbook.Worksheets[0].Comments[idx].Note = "Comment";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B5" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B5", result.Message);
        Assert.Contains("sheet", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesComment()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add("A1");
        sheet.Comments[idx].Note = "Comment to delete";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(sheet.Comments["A1"]);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("B2")]
    [InlineData("Z10")]
    public void Execute_DeletesCommentFromVariousCells(string cell)
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var idx = sheet.Comments.Add(cell);
        sheet.Comments[idx].Note = "Comment";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", cell }
        });

        _handler.Execute(context, parameters);

        Assert.Null(sheet.Comments[cell]);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_DeletesFromSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var idx = workbook.Worksheets[1].Comments.Add("A1");
        workbook.Worksheets[1].Comments[idx].Note = "Comment on Sheet2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("sheet 1", result.Message);
        Assert.Null(workbook.Worksheets[1].Comments["A1"]);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_DeletesFromFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        var idx = workbook.Worksheets[0].Comments.Add("A1");
        workbook.Worksheets[0].Comments[idx].Note = "Comment on first sheet";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("sheet 0", result.Message);
        Assert.Null(workbook.Worksheets[0].Comments["A1"]);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "" }
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
            { "cell", "InvalidCell" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
