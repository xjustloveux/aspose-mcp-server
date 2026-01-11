using AsposeMcpServer.Handlers.Pdf.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Table;

public class AddPdfTableHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Table Operations

    [SkippableFact]
    public void Execute_AddsTable()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 3 },
            { "columns", 4 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("3 rows", result);
        Assert.Contains("4 columns", result);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_AddsTableWithData()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var data = new[]
        {
            new[] { "A1", "B1" },
            new[] { "A2", "B2" }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 },
            { "data", data }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    [SkippableFact]
    public void Execute_AddsTableWithCustomPosition()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 },
            { "x", 200.0 },
            { "y", 500.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    [SkippableFact]
    public void Execute_AddsTableWithColumnWidths()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 3 },
            { "columnWidths", "100 150 200" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    [Fact]
    public void Execute_WithZeroRows_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 0 },
            { "columns", 2 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithZeroColumns_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 999 },
            { "rows", 2 },
            { "columns", 2 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
