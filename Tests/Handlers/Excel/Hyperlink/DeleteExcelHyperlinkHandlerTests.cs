using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Hyperlink;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Hyperlink;

public class DeleteExcelHyperlinkHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Various Hyperlink Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesVariousIndices(int index)
    {
        var workbook = CreateWorkbookWithHyperlinks(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", index }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_DeletesFromCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Hyperlinks.Add("A1", 1, 1, "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "cell", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(workbook.Worksheets[1].Hyperlinks);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesHyperlinkByCell()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeletesHyperlinkByIndex()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
    }

    [Fact]
    public void Execute_DecreasesHyperlinkCount()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var sheet = workbook.Worksheets[0];
        var initialCount = sheet.Hyperlinks.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, sheet.Hyperlinks.Count);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var workbook = CreateWorkbookWithHyperlinks(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("2 hyperlinks remaining", result.Message);
    }

    [Fact]
    public void Execute_ReturnsCellReference()
    {
        var workbook = CreateWorkbookWithHyperlink("B2", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B2", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCellOrIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

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
            { "hyperlinkIndex", 99 }
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
            { "cell", "Z99" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("No hyperlink found", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithHyperlink("A1", "https://example.com");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
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

    private static Workbook CreateWorkbookWithHyperlinks(int count)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        for (var i = 0; i < count; i++) sheet.Hyperlinks.Add($"A{i + 1}", 1, 1, $"https://example{i}.com");
        return workbook;
    }

    #endregion
}
