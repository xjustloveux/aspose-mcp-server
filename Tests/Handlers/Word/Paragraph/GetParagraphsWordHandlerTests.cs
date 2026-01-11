using System.Text.Json;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var paragraphs = json.RootElement.GetProperty("paragraphs");
        Assert.True(paragraphs.GetArrayLength() > 0);
        var firstPara = paragraphs[0];
        Assert.True(firstPara.TryGetProperty("index", out _));
        Assert.True(firstPara.TryGetProperty("location", out _));
        Assert.True(firstPara.TryGetProperty("text", out _));
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

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsParagraphsInfo()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("paragraphs", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() >= 3);
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("filters", out var filters));
        Assert.False(filters.GetProperty("includeEmpty").GetBoolean());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("filters", out var filters));
        Assert.Equal("Normal", filters.GetProperty("styleFilter").GetString());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("filters", out var filters));
        Assert.Equal(0, filters.GetProperty("sectionIndex").GetInt32());
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
