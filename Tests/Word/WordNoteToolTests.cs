using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordNoteToolTests : WordTestBase
{
    private readonly WordNoteTool _tool = new();

    [Fact]
    public async Task AddFootnote_ShouldAddFootnote()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_footnote.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_footnote_output.docx");
        var arguments = CreateArguments("add_footnote", docPath, outputPath);
        arguments["text"] = "This is a footnote";
        arguments["paragraphIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        Assert.True(footnotes.Count > 0, "Document should contain at least one footnote");
    }

    [Fact]
    public async Task AddEndnote_ShouldAddEndnote()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_endnote.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_endnote_output.docx");
        var arguments = CreateArguments("add_endnote", docPath, outputPath);
        arguments["text"] = "This is an endnote";
        arguments["paragraphIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var endnotes = doc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.True(endnotes.Count > 0, "Document should contain at least one endnote");
    }

    [Fact]
    public async Task GetFootnotes_ShouldReturnAllFootnotes()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_footnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Test footnote");
        doc.Save(docPath);

        var arguments = CreateArguments("get_footnotes", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Footnote", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteFootnote_ShouldDeleteFootnote()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Note to delete");
        doc.Save(docPath);

        var footnotesBefore = doc.GetChildNodes(NodeType.Footnote, true).Count;
        Assert.True(footnotesBefore > 0, "Footnote should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_footnote_output.docx");
        var arguments = CreateArguments("delete_footnote", docPath, outputPath);
        arguments["noteIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var footnotesAfter = resultDoc.GetChildNodes(NodeType.Footnote, true).Count;
        Assert.True(footnotesAfter < footnotesBefore,
            $"Footnote should be deleted. Before: {footnotesBefore}, After: {footnotesAfter}");
    }

    [Fact]
    public async Task EditFootnote_ShouldUpdateFootnoteText()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_footnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_footnote_output.docx");
        var arguments = CreateArguments("edit_footnote", docPath, outputPath);
        arguments["noteIndex"] = 0;
        arguments["text"] = "Updated footnote text";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("edited successfully", result);
        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
        Assert.Contains("Updated footnote text", footnotes[0].ToString(SaveFormat.Text));
    }

    [Fact]
    public async Task GetEndnotes_ShouldReturnAllEndnotes()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_endnotes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Test endnote");
        doc.Save(docPath);

        var arguments = CreateArguments("get_endnotes", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("Endnote", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("noteIndex", result);
    }

    [Fact]
    public async Task AddFootnote_WithReferenceText_ShouldInsertAtCorrectPosition()
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
        var arguments = CreateArguments("add_footnote", docPath, outputPath);
        arguments["text"] = "Footnote for target";
        arguments["referenceText"] = "target";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("added successfully", result);
        var resultDoc = new Document(outputPath);
        var footnotes = resultDoc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
        Assert.Single(footnotes);
    }

    [Fact]
    public async Task DeleteEndnote_ShouldDeleteEndnote()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote to delete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_endnote_output.docx");
        var arguments = CreateArguments("delete_endnote", docPath, outputPath);
        arguments["noteIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Deleted", result);
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Empty(endnotes);
    }

    [Fact]
    public async Task EditEndnote_ShouldUpdateEndnoteText()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_endnote.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_endnote_output.docx");
        var arguments = CreateArguments("edit_endnote", docPath, outputPath);
        arguments["noteIndex"] = 0;
        arguments["text"] = "Updated endnote text";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("edited successfully", result);
        var resultDoc = new Document(outputPath);
        var endnotes = resultDoc.GetChildNodes(NodeType.Footnote, true)
            .Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();
        Assert.Single(endnotes);
        Assert.Contains("Updated endnote text", endnotes[0].ToString(SaveFormat.Text));
    }
}