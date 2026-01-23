using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Styles;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordStyleTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordStyleToolTests : WordTestBase
{
    private readonly WordStyleTool _tool;

    public WordStyleToolTests()
    {
        _tool = new WordStyleTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void GetStyles_ShouldReturnAllStylesFromFile()
    {
        var docPath = CreateWordDocument("test_get_styles.docx");
        var result = _tool.Execute("get_styles", docPath);
        var data = GetResultData<GetWordStylesResult>(result);
        Assert.True(data.Count > 0);
        Assert.NotNull(data.ParagraphStyles);
        Assert.Contains(data.ParagraphStyles, s => s.Name == "Normal");
    }

    [Fact]
    public void CreateStyle_ShouldCreateNewStyleAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_create_style.docx");
        var outputPath = CreateTestFilePath("test_create_style_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: "CustomStyle", styleType: "paragraph", fontSize: 14, bold: true);
        var doc = new Document(outputPath);
        var style = doc.Styles["CustomStyle"];
        Assert.NotNull(style);
        Assert.Equal(14, style.Font.Size);
        Assert.True(style.Font.Bold);
    }

    [Fact]
    public void ApplyStyle_ShouldApplyStyleAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_apply_style.docx", "Test");
        var outputPath = CreateTestFilePath("test_apply_style_output.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "TestStyle");
        customStyle.Font.Size = 16;
        doc.Save(docPath);
        _tool.Execute("apply_style", docPath, outputPath: outputPath,
            styleName: "TestStyle", paragraphIndex: 0);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ApplyStyle_ToTable_ShouldApplyTableStyleAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_apply_style_table.docx");
        var outputPath = CreateTestFilePath("test_apply_style_table_output.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();
        var tableStyle = doc.Styles.Add(StyleType.Table, "TestTableStyle");
        tableStyle.Font.Size = 12;
        doc.Save(docPath);
        _tool.Execute("apply_style", docPath, outputPath: outputPath,
            styleName: "TestTableStyle", tableIndex: 0);
        var resultDoc = new Document(outputPath);
        var tables = resultDoc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Equal("TestTableStyle", tables[0].StyleName);
    }

    [Fact]
    public void CopyStyles_ShouldCopyStylesAndPersistToFile()
    {
        var sourcePath = CreateWordDocument("test_copy_styles_source.docx");
        var doc = new Document(sourcePath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "SourceStyle");
        customStyle.Font.Size = 16;
        doc.Save(sourcePath);

        var targetPath = CreateWordDocument("test_copy_styles_target.docx");
        var outputPath = CreateTestFilePath("test_copy_styles_output.docx");
        _tool.Execute("copy_styles", targetPath, outputPath: outputPath, sourceDocument: sourcePath);
        var resultDoc = new Document(outputPath);
        var copiedStyle = resultDoc.Styles["SourceStyle"];
        Assert.NotNull(copiedStyle);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_STYLES")]
    [InlineData("Get_Styles")]
    [InlineData("get_styles")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation.GetHashCode()}_case.docx");
        var result = _tool.Execute(operation, docPath);
        var data = GetResultData<GetWordStylesResult>(result);
        Assert.NotNull(data.ParagraphStyles);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_unknown_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_styles"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetStyles_WithSessionId_ShouldReturnStyles()
    {
        var docPath = CreateWordDocument("test_session_get_styles.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "SessionStyle");
        customStyle.Font.Size = 16;
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_styles", sessionId: sessionId);
        var data = GetResultData<GetWordStylesResult>(result);
        Assert.Contains(data.ParagraphStyles, s => s.Name == "SessionStyle");
        var output = GetResultOutput<GetWordStylesResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void CreateStyle_WithSessionId_ShouldCreateStyleInMemory()
    {
        var docPath = CreateWordDocument("test_session_create_style.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("create_style", sessionId: sessionId,
            styleName: "SessionCreatedStyle", styleType: "paragraph", fontSize: 20, bold: true);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("SessionCreatedStyle", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var style = doc.Styles["SessionCreatedStyle"];
        Assert.NotNull(style);
        Assert.Equal(20, style.Font.Size);
    }

    [Fact]
    public void ApplyStyle_WithSessionId_ShouldApplyStyleInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_apply_style.docx", "Test content");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "ApplySessionStyle");
        customStyle.Font.Size = 18;
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("apply_style", sessionId: sessionId,
            styleName: "ApplySessionStyle", paragraphIndex: 0);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = sessionDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        Assert.Equal("ApplySessionStyle", paragraphs[0].ParagraphFormat.StyleName);
    }

    [Fact]
    public void CopyStyles_WithSessionId_ShouldCopyToMemory()
    {
        var sourcePath = CreateWordDocument("test_copy_session_source.docx");
        var sourceDoc = new Document(sourcePath);
        sourceDoc.Styles.Add(StyleType.Paragraph, "SourceSessionStyle");
        sourceDoc.Save(sourcePath);

        var targetPath = CreateWordDocument("test_copy_session_target.docx");
        var sessionId = OpenSession(targetPath);

        var result = _tool.Execute("copy_styles", sessionId: sessionId, sourceDocument: sourcePath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Copied", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc.Styles["SourceSessionStyle"]);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_styles", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_style_path.docx");
        var doc1 = new Document(docPath1);
        doc1.Styles.Add(StyleType.Paragraph, "PathStyle");
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_style_session.docx");
        var doc2 = new Document(docPath2);
        doc2.Styles.Add(StyleType.Paragraph, "SessionStyle");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get_styles", docPath1, sessionId);
        var data = GetResultData<GetWordStylesResult>(result);

        Assert.Contains(data.ParagraphStyles, s => s.Name == "SessionStyle");
        Assert.DoesNotContain(data.ParagraphStyles, s => s.Name == "PathStyle");
    }

    #endregion
}
