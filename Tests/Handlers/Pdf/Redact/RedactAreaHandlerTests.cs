using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Redact;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Redact;

public class RedactAreaHandlerTests : PdfHandlerTestBase
{
    private readonly RedactAreaHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Area()
    {
        Assert.Equal("area", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreatePdfWithText(string text)
    {
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment(text));
        return document;
    }

    #endregion

    #region Basic Redact Area Operations

    [SkippableFact]
    public void Execute_RedactsArea()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithText("Confidential information");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "x", 50.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 50.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("redaction applied", result.ToLower());
        Assert.Contains("page 1", result.ToLower());
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithCustomFillColor()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithText("Secret data");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "x", 100.0 },
            { "y", 600.0 },
            { "width", 150.0 },
            { "height", 30.0 },
            { "fillColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("redaction applied", result.ToLower());
    }

    [SkippableFact]
    public void Execute_WithOverlayText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithText("Private information");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "x", 100.0 },
            { "y", 500.0 },
            { "width", 200.0 },
            { "height", 40.0 },
            { "overlayText", "REDACTED" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("redaction applied", result.ToLower());
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 999 },
            { "x", 50.0 },
            { "y", 700.0 },
            { "width", 100.0 },
            { "height", 50.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithPageIndexZero_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 },
            { "x", 50.0 },
            { "y", 700.0 },
            { "width", 100.0 },
            { "height", 50.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
