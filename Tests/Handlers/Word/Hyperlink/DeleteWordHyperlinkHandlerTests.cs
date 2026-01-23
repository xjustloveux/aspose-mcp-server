using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Hyperlink;

public class DeleteWordHyperlinkHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Default Hyperlink Index

    [Fact]
    public void Execute_WithoutHyperlinkIndex_DeletesFirstHyperlink()
    {
        var doc = CreateDocumentWithHyperlinks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#0", result.Message);
        Assert.Contains("deleted successfully", result.Message);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithHyperlinks(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++)
        {
            builder.InsertHyperlink($"Link {i}", $"https://example{i}.com", false);
            builder.InsertParagraph();
        }

        return doc;
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesHyperlink()
    {
        var doc = CreateDocumentWithHyperlinks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted successfully", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDeletedIndex()
    {
        var doc = CreateDocumentWithHyperlinks(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#1", result.Message);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var doc = CreateDocumentWithHyperlinks(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Remaining hyperlinks", result.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesAtVariousIndices(int index)
    {
        var doc = CreateDocumentWithHyperlinks(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", index }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted successfully", result.Message);
    }

    #endregion

    #region Keep Text Option

    [Fact]
    public void Execute_WithKeepTextTrue_UnlinksHyperlink()
    {
        var doc = CreateDocumentWithHyperlinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "keepText", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("unlinked", result.Message);
    }

    [Fact]
    public void Execute_WithKeepTextFalse_RemovesHyperlink()
    {
        var doc = CreateDocumentWithHyperlinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "keepText", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message);
    }

    [Fact]
    public void Execute_DefaultKeepTextIsFalse()
    {
        var doc = CreateDocumentWithHyperlinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("No (removed)", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidHyperlinkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithHyperlinks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeHyperlinkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithHyperlinks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_NoHyperlinks_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("No hyperlinks here");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("no hyperlinks", ex.Message);
    }

    #endregion
}
