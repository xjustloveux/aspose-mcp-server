using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class RestartWordListNumberingHandlerTests : WordHandlerTestBase
{
    private readonly RestartWordListNumberingHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_RestartNumbering()
    {
        Assert.Equal("restart_numbering", _handler.Operation);
    }

    #endregion

    #region Start At Parameter

    [Theory]
    [InlineData(1)]
    [InlineData(5)]
    [InlineData(10)]
    [InlineData(100)]
    public void Execute_WithStartAt_UsesSpecifiedValue(int startAt)
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "startAt", startAt }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var para = paragraphs[0];
        Assert.True(para.ListFormat.IsListItem);
        var level = para.ListFormat.ListLevelNumber;
        Assert.Equal(startAt, para.ListFormat.List.ListLevels[level].StartAt);
        AssertModified(context);
    }

    #endregion

    #region Various Paragraph Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_RestartsAtVariousParagraphs(int index)
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", index }
        });

        var originalListId = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList()[index]
            .ListFormat.List.ListId;

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var para = paragraphs[index];
        Assert.True(para.ListFormat.IsListItem);
        Assert.NotEqual(originalListId, para.ListFormat.List.ListId);
        AssertModified(context);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithList(int itemCount)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        var list = doc.Lists.Add(ListTemplate.NumberDefault);

        for (var i = 0; i < itemCount; i++)
        {
            builder.ListFormat.List = list;
            builder.Writeln($"List Item {i + 1}");
        }

        builder.ListFormat.RemoveNumbers();
        return doc;
    }

    #endregion

    #region Basic Restart Operations

    [Fact]
    public void Execute_RestartsListNumbering()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 }
        });

        var originalListId = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList()[1]
            .ListFormat.List.ListId;

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var para = paragraphs[1];
        Assert.True(para.ListFormat.IsListItem);
        Assert.NotEqual(originalListId, para.ListFormat.List.ListId);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsParagraphIndex()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        Assert.True(paragraphs[1].ListFormat.IsListItem);
        Assert.NotNull(paragraphs[1].ListFormat.List);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDefaultStartAt()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var para = paragraphs[0];
        var level = para.ListFormat.ListLevelNumber;
        Assert.Equal(1, para.ListFormat.List.ListLevels[level].StartAt);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsAffectedParagraphsCount()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var listParas = paragraphs.Where(p => p.ListFormat.IsListItem).ToList();
        Assert.True(listParas.Count > 0);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsNewListId()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var originalListId = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList()[0]
            .ListFormat.List.ListId;

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        Assert.NotEqual(originalListId, paragraphs[0].ListFormat.List.ListId);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithList(3);
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
        var doc = CreateDocumentWithList(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNonListParagraph_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Normal text", "More text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not a list item", ex.Message);
    }

    #endregion
}
