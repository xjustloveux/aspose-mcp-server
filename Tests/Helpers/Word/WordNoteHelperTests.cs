using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Word;

public class WordNoteHelperTests : WordTestBase
{
    #region UpdateNoteText Tests

    [Fact]
    public void UpdateNoteText_UpdatesContent()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var note = builder.InsertFootnote(FootnoteType.Footnote, "Original text");

        WordNoteHelper.UpdateNoteText(doc, note, "Updated text");

        Assert.Contains("Updated text", note.GetText());
    }

    #endregion

    #region GetNoteTypeName Tests

    [Fact]
    public void GetNoteTypeName_WithFootnote_ReturnsFootnote()
    {
        var result = WordNoteHelper.GetNoteTypeName(FootnoteType.Footnote);

        Assert.Equal("footnote", result);
    }

    [Fact]
    public void GetNoteTypeName_WithEndnote_ReturnsEndnote()
    {
        var result = WordNoteHelper.GetNoteTypeName(FootnoteType.Endnote);

        Assert.Equal("endnote", result);
    }

    #endregion

    #region GetNotesFromDoc Tests

    [Fact]
    public void GetNotesFromDoc_WithNoNotes_ReturnsEmptyList()
    {
        var doc = new Document();

        var result = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        Assert.Empty(result);
    }

    [Fact]
    public void GetNotesFromDoc_WithFootnotes_ReturnsOnlyFootnotes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with footnote");
        builder.InsertFootnote(FootnoteType.Footnote, "This is a footnote");
        builder.Write(" and endnote");
        builder.InsertFootnote(FootnoteType.Endnote, "This is an endnote");

        var result = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        Assert.Single(result);
        Assert.Equal(FootnoteType.Footnote, result[0].FootnoteType);
    }

    [Fact]
    public void GetNotesFromDoc_WithEndnotes_ReturnsOnlyEndnotes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote");
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote");

        var result = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Endnote);

        Assert.Single(result);
        Assert.Equal(FootnoteType.Endnote, result[0].FootnoteType);
    }

    #endregion

    #region FindNote Tests

    [Fact]
    public void FindNote_WithEmptyList_ReturnsNull()
    {
        var notes = new List<Footnote>();

        var result = WordNoteHelper.FindNote(notes, null, null);

        Assert.Null(result);
    }

    [Fact]
    public void FindNote_WithNoParameters_ReturnsFirstNote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertFootnote(FootnoteType.Footnote, "First");
        builder.InsertFootnote(FootnoteType.Footnote, "Second");
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var result = WordNoteHelper.FindNote(notes, null, null);

        Assert.NotNull(result);
    }

    [Fact]
    public void FindNote_WithValidIndex_ReturnsNoteAtIndex()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertFootnote(FootnoteType.Footnote, "First");
        builder.InsertFootnote(FootnoteType.Footnote, "Second");
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var result = WordNoteHelper.FindNote(notes, null, 1);

        Assert.NotNull(result);
    }

    [Fact]
    public void FindNote_WithInvalidIndex_ReturnsNull()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertFootnote(FootnoteType.Footnote, "First");
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var result = WordNoteHelper.FindNote(notes, null, 10);

        Assert.Null(result);
    }

    [Fact]
    public void FindNote_WithNegativeIndex_ReturnsNull()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertFootnote(FootnoteType.Footnote, "First");
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var result = WordNoteHelper.FindNote(notes, null, -1);

        Assert.Null(result);
    }

    [Fact]
    public void FindNote_WithReferenceMark_ReturnsMatchingNote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var note = builder.InsertFootnote(FootnoteType.Footnote, "Custom note");
        note.ReferenceMark = "X";
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var result = WordNoteHelper.FindNote(notes, "X", null);

        Assert.NotNull(result);
        Assert.Equal("X", result.ReferenceMark);
    }

    [Fact]
    public void FindNote_WithNonExistentReferenceMark_ReturnsNull()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertFootnote(FootnoteType.Footnote, "Note");
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var result = WordNoteHelper.FindNote(notes, "NonExistent", null);

        Assert.Null(result);
    }

    #endregion

    #region InsertNoteAtDocumentEnd Tests

    [Fact]
    public void InsertNoteAtDocumentEnd_WithFootnote_InsertsFootnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Document text");

        var result = WordNoteHelper.InsertNoteAtDocumentEnd(builder, FootnoteType.Footnote, "Note text", null);

        Assert.NotNull(result);
        Assert.Equal(FootnoteType.Footnote, result.FootnoteType);
    }

    [Fact]
    public void InsertNoteAtDocumentEnd_WithEndnote_InsertsEndnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Document text");

        var result = WordNoteHelper.InsertNoteAtDocumentEnd(builder, FootnoteType.Endnote, "Endnote text", null);

        Assert.NotNull(result);
        Assert.Equal(FootnoteType.Endnote, result.FootnoteType);
    }

    [Fact]
    public void InsertNoteAtDocumentEnd_WithCustomMark_SetsReferenceMark()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Document text");

        var result = WordNoteHelper.InsertNoteAtDocumentEnd(builder, FootnoteType.Footnote, "Note", "★");

        Assert.Equal("★", result.ReferenceMark);
    }

    [Fact]
    public void InsertNoteAtDocumentEnd_WithNullCustomMark_DoesNotSetReferenceMark()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Document text");

        var result = WordNoteHelper.InsertNoteAtDocumentEnd(builder, FootnoteType.Footnote, "Note", null);

        Assert.NotNull(result);
    }

    #endregion

    #region InsertNoteAtParagraph Tests

    [Fact]
    public void InsertNoteAtParagraph_WithValidIndex_InsertsNote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph");
        builder.Writeln("Second paragraph");
        var section = doc.Sections[0];

        var result = WordNoteHelper.InsertNoteAtParagraph(builder, section, 0, FootnoteType.Footnote, "Note", null);

        Assert.NotNull(result);
    }

    [Fact]
    public void InsertNoteAtParagraph_WithMinusOne_InsertsAtEnd()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph");
        builder.Writeln("Second paragraph");
        var section = doc.Sections[0];

        var result = WordNoteHelper.InsertNoteAtParagraph(builder, section, -1, FootnoteType.Footnote, "Note", null);

        Assert.NotNull(result);
    }

    [Fact]
    public void InsertNoteAtParagraph_WithInvalidIndex_ThrowsArgumentException()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph");
        var section = doc.Sections[0];

        var ex = Assert.Throws<ArgumentException>(() =>
            WordNoteHelper.InsertNoteAtParagraph(builder, section, 100, FootnoteType.Footnote, "Note", null));

        Assert.Contains("paragraphIndex must be between", ex.Message);
    }

    #endregion

    #region InsertNoteAtReferenceText Tests

    [Fact]
    public void InsertNoteAtReferenceText_WithExistingText_InsertsNote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("This is reference text in document");

        var result = WordNoteHelper.InsertNoteAtReferenceText(
            doc, builder, "reference text", FootnoteType.Footnote, "Note content", null);

        Assert.NotNull(result);
        Assert.Equal(FootnoteType.Footnote, result.FootnoteType);
    }

    [Fact]
    public void InsertNoteAtReferenceText_WithNonExistingText_ThrowsArgumentException()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Some document text");

        var ex = Assert.Throws<ArgumentException>(() =>
            WordNoteHelper.InsertNoteAtReferenceText(
                doc, builder, "nonexistent", FootnoteType.Footnote, "Note", null));

        Assert.Contains("not found", ex.Message);
    }

    #endregion
}
