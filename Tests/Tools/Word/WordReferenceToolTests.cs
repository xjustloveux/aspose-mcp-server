using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordReferenceToolTests : WordTestBase
{
    private readonly WordReferenceTool _tool;

    public WordReferenceToolTests()
    {
        _tool = new WordReferenceTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddTableOfContents_ShouldAddTOC()
    {
        var docPath = CreateWordDocumentWithContent("test_add_toc.docx", "Heading 1\nContent\nHeading 2\nMore content");
        var doc = new Document(docPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        if (paragraphs.Count > 0) paragraphs[0].ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        if (paragraphs.Count > 2) paragraphs[2].ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_toc_output.docx");
        _tool.Execute("add_table_of_contents", docPath, outputPath: outputPath,
            title: "Table of Contents", maxLevel: 3);
        var resultDoc = new Document(outputPath);
        Assert.NotNull(resultDoc);
    }

    [Fact]
    public void UpdateTableOfContents_ShouldUpdateTOC()
    {
        var docPath = CreateWordDocument("test_update_toc.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("New Heading");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_toc_output.docx");
        _tool.Execute("update_table_of_contents", docPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public void AddCrossReference_WithReferenceType_ShouldUseReferenceType()
    {
        var docPath = CreateWordDocument("test_cross_reference.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Chapter 1");
        builder.StartBookmark("Chapter1");
        builder.Writeln("Content of Chapter 1");
        builder.EndBookmark("Chapter1");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_cross_reference_output.docx");
        var result = _tool.Execute("add_cross_reference", docPath, outputPath: outputPath,
            referenceType: "Bookmark", targetName: "Chapter1", referenceText: "See ");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Bookmark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddTableOfContents_AtEnd_ShouldInsertAtEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_toc_end.docx", "Content before TOC");
        var outputPath = CreateTestFilePath("test_toc_end_output.docx");
        var result = _tool.Execute("add_table_of_contents", docPath, outputPath: outputPath,
            position: "end", title: "Contents");
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Table of contents added", result);
    }

    [Fact]
    public void UpdateTableOfContents_WhenNoTOC_ShouldReturnMessage()
    {
        var docPath = CreateWordDocumentWithContent("test_no_toc.docx", "Document without TOC");
        var result = _tool.Execute("update_table_of_contents", docPath);
        Assert.Contains("No table of contents fields found", result);
        Assert.Contains("add_table_of_contents", result);
    }

    [Fact]
    public void UpdateTableOfContents_WithInvalidTocIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_toc_invalid_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\"");
        doc.Save(docPath);
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("update_table_of_contents", docPath, tocIndex: 5));
        Assert.Contains("tocIndex must be between", exception.Message);
    }

    [Fact]
    public void AddIndex_ShouldAddIndexEntries()
    {
        var docPath = CreateWordDocumentWithContent("test_add_index.docx", "Document content");
        var outputPath = CreateTestFilePath("test_add_index_output.docx");
        var indexEntries = new JsonArray
        {
            new JsonObject { ["text"] = "Term1" },
            new JsonObject { ["text"] = "Term2", ["subEntry"] = "SubTerm" }
        };
        var result = _tool.Execute("add_index", docPath, outputPath: outputPath,
            indexEntries: indexEntries.ToJsonString());
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Index entries added", result);
        Assert.Contains("Total entries: 2", result);
    }

    [Fact]
    public void AddIndex_WithInvalidHeadingStyle_ShouldFallbackToHeading1()
    {
        var docPath = CreateWordDocumentWithContent("test_index_invalid_style.docx", "Content");
        var outputPath = CreateTestFilePath("test_index_invalid_style_output.docx");
        var indexEntries = new JsonArray
        {
            new JsonObject { ["text"] = "TestTerm" }
        };
        var result = _tool.Execute("add_index", docPath, outputPath: outputPath,
            indexEntries: indexEntries.ToJsonString(), headingStyle: "NonExistentStyle");
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Index entries added", result);
    }

    [Fact]
    public void AddCrossReference_WithInvalidReferenceType_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_invalid_ref_type.docx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_cross_reference", docPath,
                referenceType: "InvalidType", targetName: "SomeTarget"));
        Assert.Contains("Invalid referenceType", exception.Message);
        Assert.Contains("Heading, Bookmark, Figure, Table, Equation", exception.Message);
    }

    [Fact]
    public void AddCrossReference_WithIncludeAboveBelow_ShouldAddText()
    {
        var docPath = CreateWordDocument("test_cross_ref_above.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("Target");
        builder.Writeln("Target content");
        builder.EndBookmark("Target");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_cross_ref_above_output.docx");
        var result = _tool.Execute("add_cross_reference", docPath, outputPath: outputPath,
            referenceType: "Bookmark", targetName: "Target", includeAboveBelow: true);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Cross-reference added", result);
    }

    [Fact]
    public void AddIndex_WithoutInsertingIndexField_ShouldOnlyAddEntries()
    {
        var docPath = CreateWordDocumentWithContent("test_index_no_field.docx", "Content");
        var outputPath = CreateTestFilePath("test_index_no_field_output.docx");
        var indexEntries = new JsonArray
        {
            new JsonObject { ["text"] = "Entry1" }
        };
        var result = _tool.Execute("add_index", docPath, outputPath: outputPath,
            indexEntries: indexEntries.ToJsonString(), insertIndexAtEnd: false);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Index entries added", result);
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
    public void AddCrossReference_WithMissingTargetName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_missing_target.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_cross_reference", docPath, referenceType: "Bookmark", targetName: ""));

        Assert.Contains("targetName is required", ex.Message);
    }

    [Fact]
    public void AddIndex_WithEmptyEntries_ShouldSucceedWithZeroEntries()
    {
        var docPath = CreateWordDocumentWithContent("test_empty_entries.docx", "Test content");
        var outputPath = CreateTestFilePath("test_empty_entries_output.docx");

        // Act - Empty array is valid JSON, tool should succeed with 0 entries
        var result = _tool.Execute("add_index", docPath, outputPath: outputPath, indexEntries: "[]");
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Index entries added", result);
        Assert.Contains("Total entries: 0", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddTableOfContents_WithSessionId_ShouldAddTOCInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_toc.docx", "Heading Content");
        var doc = new Document(docPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        if (paragraphs.Count > 0) paragraphs[0].ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_table_of_contents", sessionId: sessionId, title: "TOC");
        Assert.Contains("Table of contents added", result);

        // Verify in-memory document
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    [Fact]
    public void AddCrossReference_WithSessionId_ShouldAddReferenceInMemory()
    {
        var docPath = CreateWordDocument("test_session_cross_ref.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("SessionTarget");
        builder.Writeln("Target content");
        builder.EndBookmark("SessionTarget");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_cross_reference", sessionId: sessionId,
            referenceType: "Bookmark", targetName: "SessionTarget");
        Assert.Contains("Cross-reference added", result);

        // Verify in-memory document
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    [Fact]
    public void AddIndex_WithSessionId_ShouldAddIndexInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_index.docx", "Document content");
        var sessionId = OpenSession(docPath);
        var indexEntries = new JsonArray
        {
            new JsonObject { ["text"] = "SessionTerm" }
        };
        var result = _tool.Execute("add_index", sessionId: sessionId,
            indexEntries: indexEntries.ToJsonString());
        Assert.Contains("Index entries added", result);

        // Verify in-memory document
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add_table_of_contents", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_path_ref.docx", "Path document");
        var docPath2 = CreateWordDocumentWithContent("test_session_ref.docx", "Session document");

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        var result = _tool.Execute("add_table_of_contents", docPath1, sessionId, title: "Test");

        // Assert - should use sessionId
        Assert.Contains("Table of contents added", result);

        // Verify the session document was modified, not the path document
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    #endregion
}