using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class DeleteParagraphWordHandlerTests : WordHandlerTestBase
{
    private readonly DeleteParagraphWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_DeletesVariousParagraphs(int index)
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", index }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Paragraph Index

    [Fact]
    public void Execute_WithParagraphIndexMinus1_DeletesLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Last to delete");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 }
        });

        _handler.Execute(context, parameters);

        AssertDoesNotContainText(doc, "Last to delete");
        AssertModified(context);
    }

    [Fact]
    public void Execute_RemovesParagraphFromDocument()
    {
        var doc = CreateDocumentWithParagraphs("Keep this", "Delete this");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 }
        });

        _handler.Execute(context, parameters);

        AssertDoesNotContainText(doc, "Delete this");
        AssertContainsText(doc, "Keep this");
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Remaining", result);
    }

    [Fact]
    public void Execute_ReturnsContentPreview()
    {
        var doc = CreateDocumentWithParagraphs("Content to preview");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Content", result);
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
    [InlineData(-2)]
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
