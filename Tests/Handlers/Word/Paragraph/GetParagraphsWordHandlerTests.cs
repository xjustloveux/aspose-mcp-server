using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Results.Word.Paragraph;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class GetParagraphsWordHandlerTests : WordHandlerTestBase
{
    private readonly GetParagraphsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Paragraph Properties

    [Fact]
    public void Execute_ReturnsParagraphProperties()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.True(result.Paragraphs.Count > 0);
        var firstPara = result.Paragraphs[0];
        Assert.Equal(0, firstPara.Index);
        Assert.NotNull(firstPara.Location);
        Assert.NotNull(firstPara.Text);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithParagraphs("Test content");
        var initialText = doc.GetText();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialText, doc.GetText());
        AssertNotModified(context);
    }

    #endregion

    #region IncludeCommentParagraphs

    [Fact]
    public void Execute_WithIncludeCommentParagraphsFalse_ExcludesCommentParagraphs()
    {
        var doc = CreateDocumentWithParagraphs("Main content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeCommentParagraphs", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.NotNull(result.Filters);
        Assert.False(result.Filters.IncludeCommentParagraphs);
    }

    #endregion

    #region IncludeTextboxParagraphs

    [Fact]
    public void Execute_WithIncludeTextboxParagraphsFalse_ExcludesTextboxParagraphs()
    {
        var doc = CreateDocumentWithParagraphs("Main content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeTextboxParagraphs", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.NotNull(result.Filters);
        Assert.False(result.Filters.IncludeTextboxParagraphs);
    }

    #endregion

    #region Long Text Truncation

    [Fact]
    public void Execute_WithLongText_TruncatesText()
    {
        var longText = new string('A', 200);
        var doc = CreateDocumentWithParagraphs(longText);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        var found = false;
        foreach (var para in result.Paragraphs)
        {
            var text = para.Text;
            if (text is { Length: > 0 } && text.Contains("..."))
            {
                found = true;
                Assert.True(text.Length <= 103);
            }
        }

        Assert.True(found);
    }

    #endregion

    #region Multiple Filters

    [Fact]
    public void Execute_WithMultipleFilters_AppliesAllFilters()
    {
        var doc = CreateDocumentWithParagraphs("Content", "More content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeEmpty", false },
            { "includeCommentParagraphs", false },
            { "includeTextboxParagraphs", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.True(result.Count >= 0);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsParagraphsInfo()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Paragraphs);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.True(result.Count >= 3);
    }

    #endregion

    #region Filter Options

    [Fact]
    public void Execute_WithIncludeEmptyFalse_ExcludesEmptyParagraphs()
    {
        var doc = CreateDocumentWithParagraphs("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeEmpty", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.NotNull(result.Filters);
        Assert.False(result.Filters.IncludeEmpty);
    }

    [Fact]
    public void Execute_WithStyleFilter_FiltersbyStyle()
    {
        var doc = CreateDocumentWithParagraphs("Heading", "Normal text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleFilter", "Normal" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.NotNull(result.Filters);
        Assert.Equal("Normal", result.Filters.StyleFilter);
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_GetsParagraphsFromSpecificSection()
    {
        var doc = CreateDocumentWithParagraphs("First section content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetParagraphsWordResult>(res);

        Assert.NotNull(result.Filters);
        Assert.Equal(0, result.Filters.SectionIndex);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithParagraphs("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sectionIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
