using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"Start at: {startAt}", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("restarted successfully", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("List numbering restarted successfully", result.Message);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Paragraph index: 1", result.Message);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Start at: 1", result.Message);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Paragraphs affected:", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("New list ID:", result.Message);
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
