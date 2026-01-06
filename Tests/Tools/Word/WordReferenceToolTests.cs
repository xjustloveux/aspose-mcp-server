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

    #region General

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
    public void AddTableOfContents_AtEnd_ShouldInsertAtEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_toc_end.docx", "Content before TOC");
        var outputPath = CreateTestFilePath("test_toc_end_output.docx");
        var result = _tool.Execute("add_table_of_contents", docPath, outputPath: outputPath,
            position: "end", title: "Contents");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Table of contents added", result);
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
    public void UpdateTableOfContents_WhenNoTOC_ShouldReturnMessage()
    {
        var docPath = CreateWordDocumentWithContent("test_no_toc.docx", "Document without TOC");
        var result = _tool.Execute("update_table_of_contents", docPath);
        Assert.Contains("No table of contents fields found", result);
        Assert.Contains("add_table_of_contents", result);
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
        Assert.StartsWith("Index entries added", result);
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
        Assert.StartsWith("Index entries added", result);
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
        Assert.StartsWith("Index entries added", result);
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
        Assert.StartsWith("Cross-reference added", result);
    }

    [Theory]
    [InlineData("ADD_TABLE_OF_CONTENTS")]
    [InlineData("Add_Table_Of_Contents")]
    [InlineData("add_table_of_contents")]
    public void Execute_AddTocOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_toc_{operation.Replace("_", "")}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_toc_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath);
        Assert.StartsWith("Table of contents added", result);
    }

    [Theory]
    [InlineData("UPDATE_TABLE_OF_CONTENTS")]
    [InlineData("UpDaTe_TaBlE_oF_cOnTeNtS")]
    [InlineData("update_table_of_contents")]
    public void Execute_UpdateTocOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_update_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\"");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_update_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath);
        Assert.StartsWith("Updated", result);
    }

    [Theory]
    [InlineData("ADD_INDEX")]
    [InlineData("Add_Index")]
    [InlineData("add_index")]
    public void Execute_AddIndexOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_index_{operation.Replace("_", "")}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_index_{operation.Replace("_", "")}_output.docx");
        var indexEntries = "[{\"text\":\"Term\"}]";
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            indexEntries: indexEntries);
        Assert.StartsWith("Index entries added", result);
    }

    [Theory]
    [InlineData("ADD_CROSS_REFERENCE")]
    [InlineData("Add_Cross_Reference")]
    [InlineData("add_cross_reference")]
    public void Execute_AddCrossRefOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_crossref_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TestTarget");
        builder.Writeln("Content");
        builder.EndBookmark("TestTarget");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_crossref_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            referenceType: "Bookmark", targetName: "TestTarget");
        Assert.StartsWith("Cross-reference added", result);
    }

    [Theory]
    [InlineData("BOOKMARK")]
    [InlineData("Bookmark")]
    [InlineData("bookmark")]
    public void AddCrossReference_ReferenceTypeIsCaseInsensitive(string refType)
    {
        var docPath = CreateWordDocument($"test_reftype_{refType}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("CaseTarget");
        builder.Writeln("Content");
        builder.EndBookmark("CaseTarget");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_reftype_{refType}_output.docx");
        var result = _tool.Execute("add_cross_reference", docPath, outputPath: outputPath,
            referenceType: refType, targetName: "CaseTarget");
        Assert.StartsWith("Cross-reference added", result);
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
    public void UpdateTableOfContents_WithNegativeTocIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_neg_toc_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\"");
        doc.Save(docPath);

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("update_table_of_contents", docPath, tocIndex: -1));
        Assert.Contains("tocIndex must be between", ex.Message);
    }

    [Fact]
    public void AddIndex_WithEmptyEntries_ShouldSucceedWithZeroEntries()
    {
        var docPath = CreateWordDocumentWithContent("test_empty_entries.docx", "Test content");
        var outputPath = CreateTestFilePath("test_empty_entries_output.docx");

        var result = _tool.Execute("add_index", docPath, outputPath: outputPath, indexEntries: "[]");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Index entries added", result);
        Assert.Contains("Total entries: 0", result);
    }

    [Fact]
    public void AddIndex_WithNullEntries_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_null_entries.docx", "Content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_index", docPath, indexEntries: null));
        Assert.Contains("indexEntries is required", ex.Message);
    }

    [Fact]
    public void AddIndex_WithInvalidJson_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_invalid_json.docx", "Content");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("add_index", docPath, indexEntries: "not valid json"));
    }

    [Fact]
    public void AddIndex_WithJsonObject_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_json_object.docx", "Content");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("add_index", docPath, indexEntries: "{\"text\":\"Term\"}"));
    }

    [Fact]
    public void AddCrossReference_WithMissingReferenceType_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_missing_ref_type.docx", "Content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_cross_reference", docPath, referenceType: "", targetName: "Target"));
        Assert.Contains("referenceType is required", ex.Message);
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
    public void AddCrossReference_WithInvalidReferenceType_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_invalid_ref_type.docx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_cross_reference", docPath,
                referenceType: "InvalidType", targetName: "SomeTarget"));
        Assert.Contains("Invalid referenceType", exception.Message);
        Assert.Contains("Heading, Bookmark, Figure, Table, Equation", exception.Message);
    }

    #endregion

    #region Session

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
        Assert.StartsWith("Table of contents added", result);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    [Fact]
    public void UpdateTableOfContents_WithSessionId_ShouldUpdateInMemory()
    {
        var docPath = CreateWordDocument("test_session_update_toc.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\"");
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Test Heading");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("update_table_of_contents", sessionId: sessionId);
        Assert.StartsWith("Updated", result);

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
        Assert.StartsWith("Index entries added", result);

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
        Assert.StartsWith("Cross-reference added", result);

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

        Assert.StartsWith("Table of contents added", result);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    #endregion
}