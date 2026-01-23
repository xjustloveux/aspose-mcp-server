using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Note;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordNoteTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordNoteToolTests : WordTestBase
{
    private readonly WordNoteTool _tool;

    public WordNoteToolTests()
    {
        _tool = new WordNoteTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddFootnote_ShouldAddFootnoteAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_footnote_output.docx");
        _tool.Execute("add_footnote", docPath, outputPath: outputPath,
            text: "This is a footnote", paragraphIndex: 0);
        var doc = new Document(outputPath);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.True(footnotes.Count > 0);
    }

    [Fact]
    public void AddEndnote_ShouldAddEndnoteAndPersistToFile()
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
        Assert.True(endnotes.Count > 0);
    }

    [Fact]
    public void GetFootnotes_ShouldReturnFootnotesFromFile()
    {
        var docPath = CreateWordDocument("test_get_footnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Test footnote");
        doc.Save(docPath);
        var result = _tool.Execute("get_footnotes", docPath);
        var data = GetResultData<GetWordNotesResult>(result);
        Assert.Equal("footnote", data.NoteType);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Notes, n => n.Text.Contains("Test footnote"));
    }

    [Fact]
    public void GetEndnotes_ShouldReturnEndnotesFromFile()
    {
        var docPath = CreateWordDocument("test_get_endnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Test endnote");
        doc.Save(docPath);
        var result = _tool.Execute("get_endnotes", docPath);
        var data = GetResultData<GetWordNotesResult>(result);
        Assert.Equal("endnote", data.NoteType);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Notes, n => n.Text.Contains("Test endnote"));
    }

    [Fact]
    public void DeleteFootnote_ShouldDeleteFootnoteAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Note to delete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_footnote_output.docx");
        _tool.Execute("delete_footnote", docPath, outputPath: outputPath, noteIndex: 0);
        var resultDoc = new Document(outputPath);
        var footnotesAfter = resultDoc.GetChildNodes(NodeType.Footnote, true).Count;
        Assert.Equal(0, footnotesAfter);
    }

    [Fact]
    public void DeleteEndnote_ShouldDeleteEndnoteAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote to delete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_endnote_output.docx");
        _tool.Execute("delete_endnote", docPath, outputPath: outputPath, noteIndex: 0);
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Empty(endnotes);
    }

    [Fact]
    public void EditFootnote_ShouldUpdateFootnoteAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_edit_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_footnote_output.docx");
        _tool.Execute("edit_footnote", docPath, outputPath: outputPath,
            noteIndex: 0, text: "Updated footnote text");
        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Contains("Updated footnote text", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void EditEndnote_ShouldUpdateEndnoteAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_edit_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_endnote_output.docx");
        _tool.Execute("edit_endnote", docPath, outputPath: outputPath,
            noteIndex: 0, text: "Updated endnote text");
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Contains("Updated endnote text", endnotes[0].ToString(SaveFormat.Text));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD_FOOTNOTE")]
    [InlineData("Add_Footnote")]
    [InlineData("add_footnote")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation.Replace("_", "")}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.docx");
        _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "Case test footnote", paragraphIndex: 0);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_footnotes"));
    }

    #endregion

    #region Session Management

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
        var data = GetResultData<GetWordNotesResult>(result);
        Assert.Equal("footnote", data.NoteType);
        Assert.Contains(data.Notes, n => n.Text.Contains("Session footnote"));
        var output = GetResultOutput<GetWordNotesResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void GetEndnotes_WithSessionId_ShouldReturnEndnotes()
    {
        var docPath = CreateWordDocument("test_session_get_endnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Session endnote");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_endnotes", sessionId: sessionId);
        var data = GetResultData<GetWordNotesResult>(result);
        Assert.Equal("endnote", data.NoteType);
        Assert.Contains(data.Notes, n => n.Text.Contains("Session endnote"));
        var output = GetResultOutput<GetWordNotesResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void AddFootnote_WithSessionId_ShouldAddFootnoteInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_footnote.docx", "Test paragraph");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_footnote", sessionId: sessionId,
            text: "Session footnote text", paragraphIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Footnote added successfully", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var footnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true);
        Assert.True(footnotes.Count > 0);
    }

    [Fact]
    public void AddEndnote_WithSessionId_ShouldAddEndnoteInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_endnote.docx", "Test paragraph");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_endnote", sessionId: sessionId,
            text: "Session endnote text", paragraphIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Endnote added successfully", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var endnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.True(endnotes.Count > 0);
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
        var result = _tool.Execute("edit_footnote", sessionId: sessionId, noteIndex: 0, text: "Updated via session");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var footnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Contains("Updated via session", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void EditEndnote_WithSessionId_ShouldEditInMemory()
    {
        var docPath = CreateWordDocument("test_session_edit_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("edit_endnote", sessionId: sessionId, noteIndex: 0, text: "Updated via session");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var endnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Contains("Updated via session", endnotes[0].ToString(SaveFormat.Text));
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
        var result = _tool.Execute("delete_footnote", sessionId: sessionId, noteIndex: 0);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var footnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true);
        Assert.Equal(0, footnotes.Count);
    }

    [Fact]
    public void DeleteEndnote_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote to delete");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("delete_endnote", sessionId: sessionId, noteIndex: 0);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var endnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Empty(endnotes);
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
        var result = _tool.Execute("get_footnotes", docPath1, sessionId);
        var data = GetResultData<GetWordNotesResult>(result);
        Assert.Contains(data.Notes, n => n.Text.Contains("Session footnote"));
        Assert.DoesNotContain(data.Notes, n => n.Text.Contains("Path footnote"));
    }

    #endregion
}
