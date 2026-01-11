using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Annotation;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Annotation;

public class GetPdfAnnotationsHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfAnnotationsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException(int invalidPageIndex)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidPageIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var initialCount = doc.Pages[1].Annotations.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, doc.Pages[1].Annotations.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsAnnotationsInfo()
    {
        var doc = CreateDocumentWithAnnotation("Test annotation");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("annotations", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithAnnotations(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithNoAnnotations_ReturnsEmptyList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region Page Index

    [Fact]
    public void Execute_DefaultPageIndex_GetsFromFirstPage()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithPageIndex_GetsFromSpecificPage()
    {
        var doc = CreateDocumentWithPages(2);
        AddAnnotationToPage(doc, 2, "Annotation on page 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithPageIndexZero_GetsFromAllPages()
    {
        var doc = CreateDocumentWithPages(2);
        AddAnnotationToPage(doc, 1, "Annotation 1");
        AddAnnotationToPage(doc, 2, "Annotation 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region Annotation Properties

    [Fact]
    public void Execute_ReturnsAnnotationProperties()
    {
        var doc = CreateDocumentWithAnnotation("Test note");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var annotations = json.RootElement.GetProperty("annotations");
        Assert.True(annotations.GetArrayLength() > 0);
        var firstItem = annotations[0];
        Assert.True(firstItem.TryGetProperty("pageIndex", out _));
        Assert.True(firstItem.TryGetProperty("index", out _));
        Assert.True(firstItem.TryGetProperty("type", out _));
        Assert.True(firstItem.TryGetProperty("contents", out _));
    }

    [Fact]
    public void Execute_ReturnsRectInfo()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var firstAnnotation = json.RootElement.GetProperty("annotations")[0];
        var rect = firstAnnotation.GetProperty("rect");
        Assert.True(rect.TryGetProperty("x", out _));
        Assert.True(rect.TryGetProperty("y", out _));
        Assert.True(rect.TryGetProperty("width", out _));
        Assert.True(rect.TryGetProperty("height", out _));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithAnnotation(string text)
    {
        var doc = new Document();
        doc.Pages.Add();
        AddAnnotationToPage(doc, 1, text);
        return doc;
    }

    private static Document CreateDocumentWithAnnotations(int count)
    {
        var doc = new Document();
        doc.Pages.Add();
        for (var i = 0; i < count; i++)
            AddAnnotationToPage(doc, 1, $"Annotation {i + 1}");
        return doc;
    }

    private static void AddAnnotationToPage(Document doc, int pageIndex, string text)
    {
        var page = doc.Pages[pageIndex];
        var annotation = new TextAnnotation(page, new Rectangle(100, 700, 300, 750))
        {
            Title = "Test",
            Subject = "Test Subject",
            Contents = text
        };
        page.Annotations.Add(annotation);
    }

    #endregion
}
