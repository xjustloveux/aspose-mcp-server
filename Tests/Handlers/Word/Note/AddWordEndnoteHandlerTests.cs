using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class AddWordEndnoteHandlerTests : WordHandlerTestBase
{
    private readonly AddWordEndnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddEndnote()
    {
        Assert.Equal("add_endnote", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsEndnote()
    {
        var doc = CreateDocumentWithText("Sample text for endnote.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "This is an endnote" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added successfully", result.ToLower());
        Assert.Contains("This is an endnote", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndex_AddsEndnoteAtParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Endnote at paragraph" },
            { "paragraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added successfully", result.ToLower());
    }

    [Fact]
    public void Execute_WithCustomMark_AddsEndnoteWithCustomMark()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Custom mark endnote" },
            { "customMark", "†" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added successfully", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Endnote" },
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
