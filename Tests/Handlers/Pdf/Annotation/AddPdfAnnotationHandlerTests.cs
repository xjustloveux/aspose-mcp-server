using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Annotation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Annotation;

public class AddPdfAnnotationHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfAnnotationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", "Test" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Contains("page 1", result.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsAnnotation()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", "Test annotation" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Single(doc.Pages[1].Annotations);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Simple annotation")]
    [InlineData("Annotation with special chars !@#$%")]
    [InlineData("Multi-line\nannotation")]
    public void Execute_AddsAnnotationWithVariousTexts(string text)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", text }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1];
        Assert.Equal(text, annotation.Contents);
    }

    #endregion

    #region Position Parameters

    [Fact]
    public void Execute_WithCustomPosition_SetsPosition()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", "Test" },
            { "x", 200.0 },
            { "y", 500.0 }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1];
        Assert.Equal(200.0, annotation.Rect.LLX, 1);
        Assert.Equal(500.0, annotation.Rect.LLY, 1);
    }

    [Fact]
    public void Execute_WithDefaultPosition_UsesDefaults()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", "Test" }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1];
        Assert.Equal(100.0, annotation.Rect.LLX, 1);
        Assert.Equal(700.0, annotation.Rect.LLY, 1);
    }

    #endregion

    #region Multiple Pages

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_AddsToSpecificPage(int pageIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", pageIndex },
            { "text", "Annotation" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"page {pageIndex}", result.Message);
        Assert.Single(doc.Pages[pageIndex].Annotations);
    }

    [Fact]
    public void Execute_AddMultipleAnnotationsToSamePage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);

        for (var i = 1; i <= 3; i++)
        {
            var parameters = CreateParameters(new Dictionary<string, object?>
            {
                { "pageIndex", 1 },
                { "text", $"Annotation {i}" }
            });
            _handler.Execute(context, parameters);
        }

        Assert.Equal(3, doc.Pages[1].Annotations.Count);
    }

    #endregion

    #region Annotation Properties

    [Fact]
    public void Execute_CreatesTextAnnotation()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", "Test" }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1];
        Assert.Equal(AnnotationType.Text, annotation.AnnotationType);
    }

    [Fact]
    public void Execute_SetsAnnotationDefaults()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "text", "Test content" }
        });

        _handler.Execute(context, parameters);

        var annotation = doc.Pages[1].Annotations[1] as TextAnnotation;
        Assert.NotNull(annotation);
        Assert.Equal("Comment", annotation.Title);
        Assert.Equal("Annotation", annotation.Subject);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException(int invalidPageIndex)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidPageIndex },
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
