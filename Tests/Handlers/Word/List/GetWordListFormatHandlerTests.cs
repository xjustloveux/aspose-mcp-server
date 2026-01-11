using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class GetWordListFormatHandlerTests : WordHandlerTestBase
{
    private readonly GetWordListFormatHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFormat()
    {
        Assert.Equal("get_format", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithList(int itemCount)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        var list = doc.Lists.Add(ListTemplate.BulletDefault);

        for (var i = 0; i < itemCount; i++)
        {
            builder.ListFormat.List = list;
            builder.Writeln($"List Item {i + 1}");
        }

        builder.ListFormat.RemoveNumbers();
        return doc;
    }

    #endregion

    #region Get All List Paragraphs

    [Fact]
    public void Execute_WithNoListItems_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithParagraphs("Normal text", "More text");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.Contains("0", result);
    }

    [Fact]
    public void Execute_WithListItems_ReturnsListInfo()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.Contains("listParagraphs", result);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var doc = CreateDocumentWithList(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
    }

    #endregion

    #region Get Specific Paragraph

    [Fact]
    public void Execute_WithParagraphIndex_ReturnsSingleParagraphInfo()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("paragraphIndex", result);
    }

    [Fact]
    public void Execute_WithNonListParagraph_ReturnsNotListItemInfo()
    {
        var doc = CreateDocumentWithParagraphs("Normal text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("isListItem", result);
        Assert.Contains("false", result);
    }

    [Fact]
    public void Execute_WithListParagraph_ReturnsIsListItemTrue()
    {
        var doc = CreateDocumentWithList(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("isListItem", result);
        Assert.Contains("true", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
