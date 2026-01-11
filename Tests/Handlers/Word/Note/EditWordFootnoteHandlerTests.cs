using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class EditWordFootnoteHandlerTests : WordHandlerTestBase
{
    private readonly EditWordFootnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditFootnote()
    {
        Assert.Equal("edit_footnote", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsFootnote()
    {
        var doc = CreateDocumentWithFootnote();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated footnote text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited successfully", result.ToLower());
        Assert.Contains("Updated footnote text", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoteIndex_EditsSpecificFootnote()
    {
        var doc = CreateDocumentWithMultipleFootnotes();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "noteIndex", 1 },
            { "text", "Edited second footnote" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited successfully", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFootnote();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoFootnotes_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("No footnotes here.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New footnote text" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFootnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with footnote");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote");
        return doc;
    }

    private static Document CreateDocumentWithMultipleFootnotes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote");
        builder.Write(" more text");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote");
        return doc;
    }

    #endregion
}
