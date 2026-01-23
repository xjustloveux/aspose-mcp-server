using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class AddWordFootnoteHandlerTests : WordHandlerTestBase
{
    private readonly AddWordFootnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddFootnote()
    {
        Assert.Equal("add_footnote", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsFootnote()
    {
        var doc = CreateDocumentWithText("Sample text for footnote.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "This is a footnote" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("added successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("This is a footnote", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndex_AddsFootnoteAtParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Footnote at paragraph" },
            { "paragraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("added successfully", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithCustomMark_AddsFootnoteWithCustomMark()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Custom mark footnote" },
            { "customMark", "*" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("added successfully", result.Message, StringComparison.OrdinalIgnoreCase);
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
            { "text", "Footnote" },
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
