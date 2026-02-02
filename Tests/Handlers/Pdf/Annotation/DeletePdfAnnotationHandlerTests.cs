using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Annotation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Annotation;

public class DeletePdfAnnotationHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfAnnotationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Delete Specific Annotation

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_DeletesAnnotationAtVariousIndices(int annotationIndex)
    {
        var doc = CreateDocumentWithAnnotations(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", annotationIndex }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesAnnotation()
    {
        var doc = CreateDocumentWithAnnotation("Test annotation");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutAnnotationIndex_DeletesAllAnnotations()
    {
        var doc = CreateDocumentWithAnnotations(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Empty(doc.Pages[1].Annotations);
        AssertModified(context);
    }

    #endregion

    #region Multiple Pages

    [Fact]
    public void Execute_DeletesFromSpecificPage()
    {
        var doc = CreateDocumentWithPages(2);
        AddAnnotationToPage(doc, 1, "Page 1 annotation");
        AddAnnotationToPage(doc, 2, "Page 2 annotation");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 },
            { "annotationIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeleteAllOnPage()
    {
        var doc = CreateDocumentWithPages(2);
        AddAnnotationToPage(doc, 1, "Page 1 annotation 1");
        AddAnnotationToPage(doc, 1, "Page 1 annotation 2");
        AddAnnotationToPage(doc, 2, "Page 2 annotation");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
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
            { "pageIndex", invalidPageIndex }
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

    [Fact]
    public void Execute_OnPageWithNoAnnotations_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("No annotations", ex.Message);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSuccessMessageForSingle()
    {
        var doc = CreateDocumentWithAnnotation("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "annotationIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSuccessMessageForAll()
    {
        var doc = CreateDocumentWithAnnotations(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Empty(doc.Pages[1].Annotations);
        AssertModified(context);
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
