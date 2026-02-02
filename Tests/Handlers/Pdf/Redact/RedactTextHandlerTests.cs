using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Redact;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Redact;

public class RedactTextHandlerTests : PdfHandlerTestBase
{
    private readonly RedactTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Text()
    {
        Assert.Equal("text", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a PDF document with text that is properly searchable.
    ///     The document is saved and reloaded to ensure TextFragmentAbsorber
    ///     can find the text content (required for redaction to work).
    /// </summary>
    private static Document CreatePdfWithTextPersisted(string text)
    {
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment(text));

        // Save to memory stream and reload to properly index text content
        // Note: Do not dispose the stream - Aspose.Pdf keeps it open for lazy loading
        var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        return new Document(stream);
    }

    #endregion

    #region Basic Redact Text Operations

    [SkippableFact]
    public void Execute_RedactsTextOccurrences()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTextPersisted("This is confidential information.");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textToRedact", "confidential" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithNoMatches_ReturnsNoOccurrences()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTextPersisted("This is a test document.");
        var initialAnnotationCount = document.Pages[1].Annotations.Count;
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textToRedact", "nonexistent" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal(initialAnnotationCount, document.Pages[1].Annotations.Count);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithCaseSensitiveFalse()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTextPersisted("SECRET and secret data.");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textToRedact", "SECRET" },
            { "caseSensitive", false }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSpecificPage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTextPersisted("Private data on page.");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textToRedact", "Private" },
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithFillColorAndOverlayText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTextPersisted("Sensitive data here.");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textToRedact", "Sensitive" },
            { "fillColor", "#000000" },
            { "overlayText", "REDACTED" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textToRedact", "test" },
            { "pageIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
