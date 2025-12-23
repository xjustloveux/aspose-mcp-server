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
        arguments["noteText"] = "This is a footnote";
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
        arguments["noteText"] = "This is an endnote";
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
}