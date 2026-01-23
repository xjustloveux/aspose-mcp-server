using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Hyperlink;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Hyperlink;

public class EditExcelHyperlinkHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Multiple Changes

    [Fact]
    public void Execute_WithMultipleChanges_UpdatesAll()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://old.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://new.com" },
            { "displayText", "New Text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("url=", result.Message);
        Assert.Contains("displayText=", result.Message);
    }

    #endregion

    #region No Changes

    [Fact]
    public void Execute_WithNoChanges_ReturnsUnchangedMessage()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("unchanged", result.Message);
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_EditsOnCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Hyperlinks.Add("A1", 1, 1, "https://old.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "cell", "A1" },
            { "url", "https://new.com" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("https://new.com", workbook.Worksheets[1].Hyperlinks[0].Address);
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

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsHyperlinkByCell()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://old.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://new.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsHyperlinkByIndex()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://old.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "url", "https://new.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message);
    }

    [Fact]
    public void Execute_UpdatesUrl()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://old.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://updated.com" }
        });

        _handler.Execute(context, parameters);

        var sheet = workbook.Worksheets[0];
        Assert.Equal("https://updated.com", sheet.Hyperlinks[0].Address);
    }

    [Fact]
    public void Execute_ReturnsUrlChange()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://old.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://new.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("url=https://new.com", result.Message);
    }

    #endregion

    #region Display Text Parameter

    [Fact]
    public void Execute_UpdatesDisplayText()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "displayText", "New Display" }
        });

        _handler.Execute(context, parameters);

        var sheet = workbook.Worksheets[0];
        Assert.Equal("New Display", sheet.Hyperlinks[0].TextToDisplay);
    }

    [Fact]
    public void Execute_ReturnsDisplayTextChange()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "displayText", "Click Me" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("displayText=Click Me", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCellOrIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "url", "https://new.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("hyperlinkIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidHyperlinkIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 99 },
            { "url", "https://new.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNonExistentCell_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "Z99" },
            { "url", "https://new.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("No hyperlink found", ex.Message);
    }

    #endregion
}
