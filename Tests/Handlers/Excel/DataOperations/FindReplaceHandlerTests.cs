using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Results.Excel.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

public class FindReplaceHandlerTests : ExcelHandlerTestBase
{
    private readonly FindReplaceHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_FindReplace()
    {
        Assert.Equal("find_replace", _handler.Operation);
    }

    #endregion

    #region Basic Find Replace Operations

    [Fact]
    public void Execute_ReplacesTextInWorksheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello World");
        workbook.Worksheets[0].Cells["A2"].PutValue("Hello Again");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Hello" },
            { "replaceText", "Hi" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<FindReplaceResult>(res);

        Assert.Equal("Hello", result.FindText);
        Assert.Equal("Hi", result.ReplaceText);
        Assert.Equal(2, result.ReplacementCount);
        Assert.Equal("Hi World", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Hi Again", workbook.Worksheets[0].Cells["A2"].StringValue);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_ReplacesInSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        workbook.Worksheets[0].Cells["A1"].PutValue("Test");
        workbook.Worksheets[1].Cells["A1"].PutValue("Test");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Test" },
            { "replaceText", "Changed" },
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<FindReplaceResult>(res);

        Assert.Equal(1, result.ReplacementCount);
        Assert.Equal("Test", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Changed", workbook.Worksheets[1].Cells["A1"].StringValue);
    }

    [Fact]
    public void Execute_WithMatchCase_RespectsCaseSensitivity()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello");
        workbook.Worksheets[0].Cells["A2"].PutValue("hello");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Hello" },
            { "replaceText", "Hi" },
            { "matchCase", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<FindReplaceResult>(res);

        Assert.Equal(1, result.ReplacementCount);
        Assert.Equal("Hi", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("hello", workbook.Worksheets[0].Cells["A2"].StringValue);
    }

    [Fact]
    public void Execute_WithMatchEntireCell_MatchesWholeContent()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Test");
        workbook.Worksheets[0].Cells["A2"].PutValue("Test Data");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Test" },
            { "replaceText", "Changed" },
            { "matchEntireCell", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<FindReplaceResult>(res);

        Assert.Equal(1, result.ReplacementCount);
        Assert.Equal("Changed", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Test Data", workbook.Worksheets[0].Cells["A2"].StringValue);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFindText_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "replaceText", "Hi" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutReplaceText_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Hello" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Result Properties

    [Fact]
    public void Execute_ReturnsCorrectProperties()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Original Text");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Original" },
            { "replaceText", "New" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<FindReplaceResult>(res);

        Assert.Equal("Original", result.FindText);
        Assert.Equal("New", result.ReplaceText);
        Assert.Equal(1, result.ReplacementCount);
    }

    [Fact]
    public void Execute_WithNoMatches_ReturnsZeroCount()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello World");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "NotFound" },
            { "replaceText", "Replaced" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<FindReplaceResult>(res);

        Assert.Equal(0, result.ReplacementCount);
    }

    #endregion
}
