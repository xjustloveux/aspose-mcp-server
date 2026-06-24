using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Word.List;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Execute_WithListItems_ReturnsListInfo()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Count >= 3);
        Assert.NotNull(result.ListParagraphs);
    }

    [Fact]
    public void Execute_ReturnsResult()
    {
        var doc = CreateDocumentWithList(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatResult>(res);

        Assert.NotNull(result);
    }

    [Fact]
    public void Execute_ListInSecondSection_ReportsSectionRelativeAddress()
    {
        // A list item living in the second section must be reported in that section's Body index
        // space (sectionIndex 1, paragraphIndex 0), not as a flat global document-order index, so
        // the address round-trips back to edit/format/delete.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Intro paragraph in section 0");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        var list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;
        builder.Writeln("List item in section 1");
        builder.ListFormat.RemoveNumbers();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatResult>(res);

        var listItem = Assert.Single(result.ListParagraphs);
        Assert.Equal("Body", listItem.StoryType);
        Assert.Equal(1, listItem.SectionIndex);
        Assert.Equal(0, listItem.ParagraphIndex);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatSingleResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.ParagraphIndex);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatSingleResult>(res);

        Assert.NotNull(result);
        Assert.False(result.IsListItem);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatSingleResult>(res);

        Assert.NotNull(result);
        Assert.True(result.IsListItem);
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
    public void Execute_WithParagraphIndexMinusOne_ReturnsLastParagraphInfo()
    {
        var doc = CreateDocumentWithList(3);
        var lastIndex = doc.GetChildNodes(NodeType.Paragraph, true).Count - 1;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordListFormatSingleResult>(res);

        Assert.NotNull(result);
        Assert.Equal(lastIndex, result.ParagraphIndex);
    }

    #endregion
}
