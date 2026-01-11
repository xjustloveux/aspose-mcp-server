using System.Text.Json;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class GetParagraphFormatWordHandlerTests : WordHandlerTestBase
{
    private readonly GetParagraphFormatWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFormat()
    {
        Assert.Equal("get_format", _handler.Operation);
    }

    #endregion

    #region Include Run Details

    [Fact]
    public void Execute_WithIncludeRunDetailsTrue_ReturnsRunInfo()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "includeRunDetails", true }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("runCount", out _));
    }

    #endregion

    #region Various Paragraph Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithVariousParagraphIndices_ReturnsCorrectParagraph(int index)
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", index }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(index, json.RootElement.GetProperty("paragraphIndex").GetInt32());
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithParagraphs("Test content");
        var initialText = doc.GetText();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialText, doc.GetText());
        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsFormatInfo()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("paragraphIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("text", out _));
        Assert.True(json.RootElement.TryGetProperty("paragraphFormat", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsParagraphFormatProperties()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var format = json.RootElement.GetProperty("paragraphFormat");
        Assert.True(format.TryGetProperty("alignment", out _));
        Assert.True(format.TryGetProperty("leftIndent", out _));
        Assert.True(format.TryGetProperty("spaceBefore", out _));
        Assert.True(format.TryGetProperty("spaceAfter", out _));
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("paragraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(100)]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
