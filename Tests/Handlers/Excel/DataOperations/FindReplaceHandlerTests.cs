using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result.ToLower());
        Assert.Contains("2", result);
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

        _handler.Execute(context, parameters);

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

        _handler.Execute(context, parameters);

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

        _handler.Execute(context, parameters);

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
}
