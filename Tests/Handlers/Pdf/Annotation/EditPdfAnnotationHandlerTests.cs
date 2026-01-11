using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Annotation;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Annotation;

public class EditPdfAnnotationHandlerTests : PdfHandlerTestBase
{
    private readonly EditPdfAnnotationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 },
            { "text", "Updated" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Annotation 1", result);
        Assert.Contains("page 1", result);
        Assert.Contains("updated", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsAnnotation()
    {
        var doc = CreateDocumentWithAnnotation("Original text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 },
            { "text", "Updated text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Updated text", doc.Pages[1].Annotations[1].Contents);
        AssertModified(context);
    }

    [Theory]
    [InlineData("New content")]
    [InlineData("Updated annotation text")]
    [InlineData("Special chars: !@#$%")]
    public void Execute_UpdatesAnnotationText(string newText)
    {
        var doc = CreateDocumentWithAnnotation("Original");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 },
            { "text", newText }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(newText, doc.Pages[1].Annotations[1].Contents);
    }

    #endregion

    #region Title and Subject

    [Fact]
    public void Execute_WithTitle_UpdatesTitle()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 },
            { "title", "New Title" }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1] as MarkupAnnotation;
        Assert.NotNull(annotation);
        Assert.Equal("New Title", annotation.Title);
    }

    [Fact]
    public void Execute_WithSubject_UpdatesSubject()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 },
            { "subject", "New Subject" }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1] as MarkupAnnotation;
        Assert.NotNull(annotation);
        Assert.Equal("New Subject", annotation.Subject);
    }

    [Fact]
    public void Execute_WithAllProperties_UpdatesAll()
    {
        var doc = CreateDocumentWithAnnotation("Original");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 },
            { "text", "Updated text" },
            { "title", "Updated Title" },
            { "subject", "Updated Subject" }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1] as MarkupAnnotation;
        Assert.NotNull(annotation);
        Assert.Equal("Updated text", annotation.Contents);
        Assert.Equal("Updated Title", annotation.Title);
        Assert.Equal("Updated Subject", annotation.Subject);
    }

    [Fact]
    public void Execute_WithoutOptionalParams_PreservesOriginal()
    {
        var doc = CreateDocumentWithAnnotation("Original text");
        var annotation = doc.Pages[1].Annotations[1] as MarkupAnnotation;
        var originalTitle = annotation?.Title;
        var originalSubject = annotation?.Subject;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 }
        });

        _handler.Execute(context, parameters);

        annotation = doc.Pages[1].Annotations[1] as MarkupAnnotation;
        Assert.Equal("Original text", annotation?.Contents);
        Assert.Equal(originalTitle, annotation?.Title);
        Assert.Equal(originalSubject, annotation?.Subject);
    }

    #endregion

    #region Multiple Annotations

    [Fact]
    public void Execute_EditsSpecificAnnotation()
    {
        var doc = CreateDocumentWithAnnotations(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 2 },
            { "text", "Updated second annotation" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Annotation 1", doc.Pages[1].Annotations[1].Contents);
        Assert.Equal("Updated second annotation", doc.Pages[1].Annotations[2].Contents);
        Assert.Equal("Annotation 3", doc.Pages[1].Annotations[3].Contents);
    }

    [Fact]
    public void Execute_OnDifferentPage_EditsCorrectAnnotation()
    {
        var doc = CreateDocumentWithPages(2);
        AddAnnotationToPage(doc, 1, "Page 1 annotation");
        AddAnnotationToPage(doc, 2, "Page 2 annotation");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 },
            { "annotationIndex", 1 },
            { "text", "Updated page 2" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Page 1 annotation", doc.Pages[1].Annotations[1].Contents);
        Assert.Equal("Updated page 2", doc.Pages[2].Annotations[1].Contents);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "annotationIndex", 1 },
            { "text", "Updated" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutAnnotationIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", "Updated" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("annotationIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException(int invalidPageIndex)
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidPageIndex },
            { "annotationIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidAnnotationIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("annotationIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
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
            Title = "Original Title",
            Subject = "Original Subject",
            Contents = text
        };
        page.Annotations.Add(annotation);
    }

    #endregion
}
