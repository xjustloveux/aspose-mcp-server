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

    #region General

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
    public void AddFootnote_WithReferenceText_ShouldInsertAtCorrectPosition()
    {
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
        Assert.StartsWith("Footnote added successfully", result);
        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
    }

    [Fact]
    public void AddFootnote_WithCustomMark_ShouldSetCustomMark()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote_custom_mark.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_footnote_custom_mark_output.docx");
        _tool.Execute("add_footnote", docPath, outputPath: outputPath,
            text: "Custom marked footnote", paragraphIndex: 0, customMark: "*");

        var doc = new Document(outputPath);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
        Assert.Equal("*", footnotes[0].ReferenceMark);
    }

    [Fact]
    public void AddFootnote_AtDocumentEnd_ShouldInsertAtEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote_end.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_footnote_end_output.docx");
        _tool.Execute("add_footnote", docPath, outputPath: outputPath,
            text: "End footnote", paragraphIndex: -1);

        var doc = new Document(outputPath);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
        Assert.Contains("End footnote", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void AddFootnote_WithoutParagraphIndex_ShouldInsertAtDocumentEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote_no_para.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_footnote_no_para_output.docx");
        var result = _tool.Execute("add_footnote", docPath, outputPath: outputPath,
            text: "Default position footnote");

        Assert.StartsWith("Footnote added successfully", result);
        var doc = new Document(outputPath);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
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
    public void AddEndnote_WithCustomMark_ShouldSetCustomMark()
    {
        var docPath = CreateWordDocumentWithContent("test_add_endnote_custom_mark.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_endnote_custom_mark_output.docx");
        _tool.Execute("add_endnote", docPath, outputPath: outputPath,
            text: "Custom marked endnote", paragraphIndex: 0, customMark: "†");

        var doc = new Document(outputPath);
        var endnotes = doc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Single(endnotes);
        Assert.Equal("†", endnotes[0].ReferenceMark);
    }

    [Fact]
    public void AddEndnote_WithReferenceText_ShouldInsertAtCorrectPosition()
    {
        var docPath = CreateTestFilePath("test_add_endnote_ref.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("The ");
        builder.Write("target");
        builder.Write(" word is here.");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_endnote_ref_output.docx");
        var result = _tool.Execute("add_endnote", docPath, outputPath: outputPath,
            text: "Endnote for target", referenceText: "target");
        Assert.StartsWith("Endnote added successfully", result);

        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Single(endnotes);
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
    public void GetFootnotes_WithNoFootnotes_ShouldReturnEmptyList()
    {
        var docPath = CreateWordDocumentWithContent("test_get_no_footnotes.docx", "No footnotes here");
        var result = _tool.Execute("get_footnotes", docPath);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("\"notes\": []", result);
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
    public void GetEndnotes_WithNoEndnotes_ShouldReturnEmptyList()
    {
        var docPath = CreateWordDocumentWithContent("test_get_no_endnotes.docx", "No endnotes here");
        var result = _tool.Execute("get_endnotes", docPath);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("\"notes\": []", result);
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
    public void DeleteFootnote_ByReferenceMark_ShouldDelete()
    {
        var docPath = CreateWordDocument("test_delete_footnote_by_mark.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        var footnote = builder.InsertFootnote(FootnoteType.Footnote, "Marked footnote");
        footnote.ReferenceMark = "X";
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_footnote_by_mark_output.docx");
        var result = _tool.Execute("delete_footnote", docPath, outputPath: outputPath, referenceMark: "X");
        Assert.StartsWith("Deleted 1 footnote", result);

        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true);
        Assert.Equal(0, footnotes.Count);
    }

    [Fact]
    public void DeleteFootnote_WithoutIndexOrMark_ShouldDeleteAll()
    {
        var docPath = CreateWordDocument("test_delete_all_footnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote 3");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_all_footnotes_output.docx");
        var result = _tool.Execute("delete_footnote", docPath, outputPath: outputPath);
        Assert.StartsWith("Deleted 3 footnote", result);

        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true);
        Assert.Equal(0, footnotes.Count);
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
        Assert.StartsWith("Deleted 1 endnote", result);
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Empty(endnotes);
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
        Assert.StartsWith("Footnote edited successfully", result);
        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
        Assert.Contains("Updated footnote text", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public void EditFootnote_ByReferenceMark_ShouldEdit()
    {
        var docPath = CreateWordDocument("test_edit_footnote_by_mark.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        var footnote = builder.InsertFootnote(FootnoteType.Footnote, "Original text");
        footnote.ReferenceMark = "Y";
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_footnote_by_mark_output.docx");
        var result = _tool.Execute("edit_footnote", docPath, outputPath: outputPath,
            referenceMark: "Y", text: "Edited by reference mark");
        Assert.StartsWith("Footnote edited successfully", result);

        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
        Assert.Contains("Edited by reference mark", footnotes[0].ToString(SaveFormat.Text));
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
        Assert.StartsWith("Endnote edited successfully", result);
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Single(endnotes);
        Assert.Contains("Updated endnote text", endnotes[0].ToString(SaveFormat.Text));
    }

    [Theory]
    [InlineData("GET_FOOTNOTES")]
    [InlineData("Get_Footnotes")]
    [InlineData("get_footnotes")]
    public void Execute_GetFootnotesOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_getfn_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Test footnote");
        doc.Save(docPath);

        var result = _tool.Execute(operation, docPath);
        Assert.Contains("footnote", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("GET_ENDNOTES")]
    [InlineData("Get_Endnotes")]
    [InlineData("get_endnotes")]
    public void Execute_GetEndnotesOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_geten_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Test endnote");
        doc.Save(docPath);

        var result = _tool.Execute(operation, docPath);
        Assert.Contains("endnote", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("ADD_FOOTNOTE")]
    [InlineData("Add_Footnote")]
    [InlineData("add_footnote")]
    public void Execute_AddFootnoteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_addfn_{operation.Replace("_", "")}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_addfn_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "Case test footnote", paragraphIndex: 0);
        Assert.StartsWith("Footnote added successfully", result);
    }

    [Theory]
    [InlineData("ADD_ENDNOTE")]
    [InlineData("Add_Endnote")]
    [InlineData("add_endnote")]
    public void Execute_AddEndnoteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_adden_{operation.Replace("_", "")}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_adden_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "Case test endnote", paragraphIndex: 0);
        Assert.StartsWith("Endnote added successfully", result);
    }

    [Theory]
    [InlineData("DELETE_FOOTNOTE")]
    [InlineData("Delete_Footnote")]
    [InlineData("delete_footnote")]
    public void Execute_DeleteFootnoteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_delfn_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "To delete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_delfn_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, noteIndex: 0);
        Assert.StartsWith("Deleted 1 footnote", result);
    }

    [Theory]
    [InlineData("DELETE_ENDNOTE")]
    [InlineData("Delete_Endnote")]
    [InlineData("delete_endnote")]
    public void Execute_DeleteEndnoteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_delen_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "To delete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_delen_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, noteIndex: 0);
        Assert.StartsWith("Deleted 1 endnote", result);
    }

    [Theory]
    [InlineData("EDIT_FOOTNOTE")]
    [InlineData("Edit_Footnote")]
    [InlineData("edit_footnote")]
    public void Execute_EditFootnoteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_editfn_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Original");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_editfn_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            noteIndex: 0, text: "Edited");
        Assert.StartsWith("Footnote edited successfully", result);
    }

    [Theory]
    [InlineData("EDIT_ENDNOTE")]
    [InlineData("Edit_Endnote")]
    [InlineData("edit_endnote")]
    public void Execute_EditEndnoteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_editen_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Original");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_editen_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            noteIndex: 0, text: "Edited");
        Assert.StartsWith("Endnote edited successfully", result);
    }

    #endregion

    #region Exception

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
    public void AddEndnote_WithMissingText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_endnote_missing_text.docx", "Test");
        var outputPath = CreateTestFilePath("test_add_endnote_missing_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_endnote", docPath, outputPath: outputPath, paragraphIndex: 0));

        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddFootnote_WithInvalidSectionIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote_invalid_section.docx", "Test");
        var outputPath = CreateTestFilePath("test_add_footnote_invalid_section_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_footnote", docPath, outputPath: outputPath,
                text: "Footnote", paragraphIndex: 0, sectionIndex: 999));

        Assert.Contains("sectionIndex", ex.Message);
    }

    [Fact]
    public void AddFootnote_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote_invalid_para.docx", "Test");
        var outputPath = CreateTestFilePath("test_add_footnote_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_footnote", docPath, outputPath: outputPath,
                text: "Footnote", paragraphIndex: 999));

        Assert.Contains("paragraphIndex", ex.Message);
    }

    [Fact]
    public void AddFootnote_WithReferenceTextNotFound_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_footnote_ref_not_found.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_footnote_ref_not_found_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_footnote", docPath, outputPath: outputPath,
                text: "Footnote", referenceText: "nonexistent text"));

        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
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
    public void DeleteEndnote_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_delete_endnote_invalid_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Single endnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_endnote_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_endnote", docPath, outputPath: outputPath, noteIndex: 999));

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

    [Fact]
    public void EditFootnote_WithMissingText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_footnote_missing_text.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Original");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_footnote_missing_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_footnote", docPath, outputPath: outputPath, noteIndex: 0));

        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditEndnote_WithMissingText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_endnote_missing_text.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Original");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_endnote_missing_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_endnote", docPath, outputPath: outputPath, noteIndex: 0));

        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditEndnote_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_endnote_invalid_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Single endnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_endnote_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_endnote", docPath, outputPath: outputPath, noteIndex: 999, text: "New text"));

        Assert.Contains("endnote not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session

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
        Assert.Contains("endnote", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddFootnote_WithSessionId_ShouldAddFootnoteInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_footnote.docx", "Test paragraph");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_footnote", sessionId: sessionId,
            text: "Session footnote text", paragraphIndex: 0);
        Assert.StartsWith("Footnote added successfully", result);

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
        Assert.StartsWith("Endnote added successfully", result);

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
        _tool.Execute("edit_footnote", sessionId: sessionId, noteIndex: 0, text: "Updated via session");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var footnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
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
        _tool.Execute("edit_endnote", sessionId: sessionId, noteIndex: 0, text: "Updated via session");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var endnotes = sessionDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Single(endnotes);
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
        _tool.Execute("delete_footnote", sessionId: sessionId, noteIndex: 0);

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
        _tool.Execute("delete_endnote", sessionId: sessionId, noteIndex: 0);

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

        Assert.Contains("Session footnote", result);
        Assert.DoesNotContain("Path footnote", result);
    }

    #endregion
}