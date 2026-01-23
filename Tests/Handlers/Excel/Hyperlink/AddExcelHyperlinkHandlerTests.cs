using AsposeMcpServer.Handlers.Excel.Hyperlink;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Hyperlink;

public class AddExcelHyperlinkHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Display Text Parameter

    [Fact]
    public void Execute_WithDisplayText_SetsDisplayText()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://example.com" },
            { "displayText", "Click Here" }
        });

        _handler.Execute(context, parameters);

        var sheet = workbook.Worksheets[0];
        Assert.Equal("Click Here", sheet.Hyperlinks[0].TextToDisplay);
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_AddsToCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "cell", "A1" },
            { "url", "https://example.com" }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[1].Hyperlinks);
        Assert.Empty(workbook.Worksheets[0].Hyperlinks);
    }

    #endregion

    #region Various Cell Positions

    [Theory]
    [InlineData("A1")]
    [InlineData("B5")]
    [InlineData("Z10")]
    [InlineData("AA1")]
    public void Execute_AddsHyperlinkToVariousCells(string cell)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", cell },
            { "url", "https://example.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(cell, result.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_ReturnsCellReference()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B2" },
            { "url", "https://test.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B2", result.Message);
    }

    [Fact]
    public void Execute_ReturnsUrl()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://www.example.org" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("https://www.example.org", result.Message);
    }

    [Fact]
    public void Execute_IncreasesHyperlinkCount()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var initialCount = sheet.Hyperlinks.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://example.com" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount + 1, sheet.Hyperlinks.Count);
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
            { "url", "https://example.com" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutUrl_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithExistingHyperlink_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Hyperlinks.Add("A1", 1, 1, "https://existing.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "url", "https://new.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("already has a hyperlink", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "cell", "A1" },
            { "url", "https://example.com" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
