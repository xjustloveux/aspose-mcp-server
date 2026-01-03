using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordNoteToolTests : WordTestBase
{
    private readonly WordNoteTool _tool;

    public WordNoteToolTests()
    {
        _tool = new WordNoteTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddFootnote_ShouldAddFootnote()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_footnote_output.docx");
        _tool.Execute("add_footnote", docPath, outputPath: outputPath,
            text: "This is a footnote", paragraphIndex: 0);
        var doc = new Document(outputPath);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.True(footnotes.Count > 0, "Document should contain at least one footnote");
        Assert.Contains("This is a footnote", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void AddEndnote_ShouldAddEndnote()
    {
        var docPath = CreateWordDocumentWithContent("test_add_endnote.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_endnote_output.docx");
        _tool.Execute("add_endnote", docPath, outputPath: outputPath,
            text: "This is an endnote", paragraphIndex: 0);
        var doc = new Document(outputPath);
        var endnotes = doc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.True(endnotes.Count > 0, "Document should contain at least one endnote");
        Assert.Contains("This is an endnote", endnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void GetFootnotes_ShouldReturnAllFootnotes()
    {
        var docPath = CreateWordDocument("test_get_footnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Test footnote");
        doc.Save(docPath);
        var result = _tool.Execute("get_footnotes", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Footnote", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteFootnote_ShouldDeleteFootnote()
    {
        var docPath = CreateWordDocument("test_delete_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Note to delete");
        doc.Save(docPath);

        var footnotesBefore = doc.GetChildNodes(NodeType.Footnote, true).Count;
        Assert.True(footnotesBefore > 0, "Footnote should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_footnote_output.docx");
        _tool.Execute("delete_footnote", docPath, outputPath: outputPath, noteIndex: 0);
        var resultDoc = new Document(outputPath);
        var footnotesAfter = resultDoc.GetChildNodes(NodeType.Footnote, true).Count;
        Assert.True(footnotesAfter < footnotesBefore,
            $"Footnote should be deleted. Before: {footnotesBefore}, After: {footnotesAfter}");
    }

    [Fact]
    public void EditFootnote_ShouldUpdateFootnoteText()
    {
        var docPath = CreateWordDocument("test_edit_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_footnote_output.docx");
        var result = _tool.Execute("edit_footnote", docPath, outputPath: outputPath,
            noteIndex: 0, text: "Updated footnote text");
        Assert.Contains("edited successfully", result);
        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
        Assert.Contains("Updated footnote text", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void GetEndnotes_ShouldReturnAllEndnotes()
    {
        var docPath = CreateWordDocument("test_get_endnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Test endnote");
        doc.Save(docPath);
        var result = _tool.Execute("get_endnotes", docPath);
        Assert.NotNull(result);
        Assert.Contains("Endnote", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("noteIndex", result);
    }

    [Fact]
    public void AddFootnote_WithReferenceText_ShouldInsertAtCorrectPosition()
    {
        // Arrange - Create document with specific searchable text
        var docPath = CreateTestFilePath("test_add_footnote_ref.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("The ");
        builder.Write("target");
        builder.Write(" word is here.");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_footnote_ref_output.docx");
        var result = _tool.Execute("add_footnote", docPath, outputPath: outputPath,
            text: "Footnote for target", referenceText: "target");
        Assert.Contains("added successfully", result);
        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
    }

    [Fact]
    public void DeleteEndnote_ShouldDeleteEndnote()
    {
        var docPath = CreateWordDocument("test_delete_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote to delete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_endnote_output.docx");
        var result = _tool.Execute("delete_endnote", docPath, outputPath: outputPath, noteIndex: 0);
        Assert.Contains("Deleted", result);
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Empty(endnotes);
    }

    [Fact]
    public void EditEndnote_ShouldUpdateEndnoteText()
    {
        var docPath = CreateWordDocument("test_edit_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_endnote_output.docx");
        var result = _tool.Execute("edit_endnote", docPath, outputPath: outputPath,
            noteIndex: 0, text: "Updated endnote text");
        Assert.Contains("edited successfully", result);
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Single(endnotes);
        Assert.Contains("Updated endnote text", endnotes[0].ToString(SaveFormat.Text));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void AddFootnote_WithMissingText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_missing_text.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_missing_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_footnote", docPath, outputPath: outputPath, paragraphIndex: 0));

        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteFootnote_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_delete_invalid_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Single footnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_footnote", docPath, outputPath: outputPath, noteIndex: 999));

        Assert.Contains("out of range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditFootnote_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_invalid_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Single footnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_footnote", docPath, outputPath: outputPath, noteIndex: 999, text: "New text"));

        Assert.Contains("footnote not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetFootnotes_WithSessionId_ShouldReturnFootnotes()
    {
        var docPath = CreateWordDocument("test_session_get_footnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Session footnote");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_footnotes", sessionId: sessionId);
        Assert.Contains("footnote", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddFootnote_WithSessionId_ShouldAddFootnoteInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_footnote.docx", "Test paragraph");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_footnote", sessionId: sessionId,
            text: "Session footnote text", paragraphIndex: 0);
        Assert.Contains("added successfully", result);

        // Verify in-memory document has the footnote
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var footnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true);
        Assert.True(footnotes.Count > 0);
    }

    [Fact]
    public void EditFootnote_WithSessionId_ShouldEditInMemory()
    {
        var docPath = CreateWordDocument("test_session_edit_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("edit_footnote", sessionId: sessionId, noteIndex: 0, text: "Updated via session");

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var footnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
        Assert.Contains("Updated via session", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void DeleteFootnote_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote to delete");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete_footnote", sessionId: sessionId, noteIndex: 0);

        // Assert - verify in-memory deletion
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var footnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true);
        Assert.Equal(0, footnotes.Count);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_footnotes", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_footnote.docx");
        var doc1 = new Document(docPath1);
        var builder1 = new DocumentBuilder(doc1);
        builder1.Write("Path doc");
        builder1.InsertFootnote(FootnoteType.Footnote, "Path footnote");
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_footnote.docx");
        var doc2 = new Document(docPath2);
        var builder2 = new DocumentBuilder(doc2);
        builder2.Write("Session doc");
        builder2.InsertFootnote(FootnoteType.Footnote, "Session footnote");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        var result = _tool.Execute("get_footnotes", docPath1, sessionId);

        // Assert - should use sessionId, returning Session footnote not Path footnote
        Assert.Contains("Session footnote", result);
        Assert.DoesNotContain("Path footnote", result);
    }

    #endregion
}