using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordReferenceTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordReferenceToolTests : WordTestBase
{
    private readonly WordReferenceTool _tool;

    public WordReferenceToolTests()
    {
        _tool = new WordReferenceTool(SessionManager);
    }

    #region File I/O Smoke Tests

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
        var result = _tool.Execute("add_table_of_contents", docPath, outputPath: outputPath,
            title: "Table of Contents", maxLevel: 3);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Table of contents added", data.Message);
        Assert.True(File.Exists(outputPath));
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
        var result = _tool.Execute("update_table_of_contents", docPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Updated", data.Message);
        Assert.True(File.Exists(outputPath));
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Index entries added", data.Message);
    }

    [Fact]
    public void AddCrossReference_ShouldAddReference()
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
        Assert.True(File.Exists(outputPath));
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Cross-reference added", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD_TABLE_OF_CONTENTS")]
    [InlineData("Add_Table_Of_Contents")]
    [InlineData("add_table_of_contents")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_toc_{operation.Replace("_", "")}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_toc_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Table of contents added", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Table of contents added", data.Message);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Index entries added", data.Message);
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
        var result = _tool.Execute("add_table_of_contents", docPath1, sessionId, title: "Test");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Table of contents added", data.Message);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    #endregion
}
