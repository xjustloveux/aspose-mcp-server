using Aspose.Words;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Text;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordTextTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordTextToolTests : WordTestBase
{
    private readonly WordTextTool _tool;

    public WordTextToolTests()
    {
        _tool = new WordTextTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddText_ShouldAddTextAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_text.docx");
        var outputPath = CreateTestFilePath("test_add_text_output.docx");

        var result = _tool.Execute("add", docPath, outputPath: outputPath, text: "Hello World");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0);
        var hasExactText = paragraphs.Any(p => p.GetText().Contains("Hello World"));
        Assert.True(hasExactText);
    }

    [Fact]
    public void ReplaceText_ShouldReplaceAndPersistToFile()
    {
        var content = "Dear Customer, Thank you for contacting Customer Support.";
        var docPath = CreateWordDocumentWithContent("test_replace_text.docx", content);
        var outputPath = CreateTestFilePath("test_replace_text_output.docx");

        _tool.Execute("replace", docPath, outputPath: outputPath, find: "Customer", replace: "Client");

        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.DoesNotContain("Customer", text);
        Assert.Contains("Client", text);
    }

    [Fact]
    public void SearchText_ShouldFindTextInFile()
    {
        var docPath = CreateWordDocumentWithContent("test_search_text.docx", "This is a test document");

        var result = _tool.Execute("search", docPath, searchText: "test");

        var data = GetResultData<TextSearchResult>(result);
        Assert.Equal(1, data.MatchCount);
        Assert.Single(data.Matches);
        Assert.Equal("test", data.SearchText);
    }

    [Fact]
    public void FormatText_ShouldApplyFormattingAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_format_text.docx", "Format this text");
        var outputPath = CreateTestFilePath("test_format_text_output.docx");

        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, bold: true, italic: true);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        if (runs.Count > 0)
        {
            Assert.True(runs[0].Font.Bold);
            Assert.True(runs[0].Font.Italic);
        }
    }

    [Fact]
    public void InsertTextAtPosition_ShouldInsertAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_insert_position.docx", "First", "Third");
        var outputPath = CreateTestFilePath("test_insert_position_output.docx");

        _tool.Execute("insert_at_position", docPath, outputPath: outputPath,
            insertParagraphIndex: 0, charIndex: 0, text: "Second ");

        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var docText = doc.GetText();
        Assert.Contains("Second", docText);
    }

    [Fact]
    public void DeleteText_ShouldDeleteAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_text.docx", "Delete this text");
        var outputPath = CreateTestFilePath("test_delete_text_output.docx");

        _tool.Execute("delete", docPath, outputPath: outputPath, searchText: "Delete ");

        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.DoesNotContain("Delete", text);
    }

    [Fact]
    public void DeleteRange_ShouldDeleteRangeAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_range.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_delete_range_output.docx");

        _tool.Execute("delete_range", docPath, outputPath: outputPath,
            startParagraphIndex: 0, startCharIndex: 0, endParagraphIndex: 1, endCharIndex: 0);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AddTextWithStyle_ShouldApplyStyleAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_text_with_style.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "CustomTextStyle");
        customStyle.Font.Size = 16;
        customStyle.Font.Bold = true;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_text_with_style_output.docx");
        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Styled Text", styleName: "CustomTextStyle");

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Styled Text"));
        Assert.NotNull(para);
        Assert.Equal("CustomTextStyle", para.ParagraphFormat.StyleName);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("AdD")]
    [InlineData("add")]
    public void Execute_OperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation}_case.docx");

        var result = _tool.Execute(operation, docPath, text: "Test");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        if (!IsEvaluationMode())
        {
            var doc = new Document(docPath);
            Assert.Contains("Test", doc.GetText());
        }
    }

    [Theory]
    [InlineData("unknown_operation")]
    [InlineData("invalid")]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException(string operation)
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(operation, docPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddText_WithSessionId_ShouldAddTextInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_text.docx");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("add", sessionId: sessionId, text: "Session Text");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var text = doc.GetText();
        Assert.Contains("Session Text", text);
    }

    [Fact]
    public void SearchText_WithSessionId_ShouldSearchInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_search.docx", "Searchable content here");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("search", sessionId: sessionId, searchText: "Searchable");

        var data = GetResultData<TextSearchResult>(result);
        Assert.Equal(1, data.MatchCount);
        Assert.Equal("Searchable", data.SearchText);
    }

    [Fact]
    public void ReplaceText_WithSessionId_ShouldReplaceInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_replace.docx", "Original text to replace");
        var sessionId = OpenSession(docPath);

        _tool.Execute("replace", sessionId: sessionId, find: "Original", replace: "Modified");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.Contains("Modified", text);
        Assert.DoesNotContain("Original", text);
    }

    [Fact]
    public void FormatText_WithSessionId_ShouldFormatInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_format.docx", "Format this text");
        var sessionId = OpenSession(docPath);

        _tool.Execute("format", sessionId: sessionId, paragraphIndex: 0, runIndex: 0, bold: true);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.True(runs.Count > 0);
        Assert.True(runs[0].Font.Bold);
    }

    [Fact]
    public void InsertAtPosition_WithSessionId_ShouldInsertInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_insert.docx", "Original text");
        var sessionId = OpenSession(docPath);

        _tool.Execute("insert_at_position", sessionId: sessionId,
            insertParagraphIndex: 0, charIndex: 0, text: "Inserted: ");

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var text = doc.GetText();
        Assert.Contains("Inserted:", text);
        Assert.Contains("Original text", text);
    }

    [Fact]
    public void DeleteText_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Keep this ");
        builder.Write("DeleteMe");
        builder.Write(" and this");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, searchText: "DeleteMe");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.DoesNotContain("DeleteMe", text);
        Assert.Contains("Keep this", text);
    }

    [Fact]
    public void AddWithStyle_WithSessionId_ShouldAddInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_style.docx");
        var sessionId = OpenSession(docPath);

        _tool.Execute("add_with_style", sessionId: sessionId,
            text: "Styled Session Text", styleName: "Normal");

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var text = doc.GetText();
        Assert.Contains("Styled Session Text", text);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("search", sessionId: "invalid_session_id", searchText: "test"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_text_path.docx", "PathContent");
        var docPath2 = CreateWordDocumentWithContent("test_text_session.docx", "SessionContent");

        var sessionId = OpenSession(docPath2);

        var result = _tool.Execute("search", docPath1, sessionId, searchText: "Content");

        var data = GetResultData<TextSearchResult>(result);
        Assert.Equal(1, data.MatchCount);
        Assert.Contains("Session", data.Matches[0].Context);
    }

    #endregion
}
