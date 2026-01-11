using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordFileTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordFileToolTests : WordTestBase
{
    private readonly WordFileTool _tool;

    public WordFileToolTests()
    {
        _tool = new WordFileTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void CreateDocument_ShouldCreateNewDocument()
    {
        var outputPath = CreateTestFilePath("test_create_document.docx");
        _tool.Execute("create", outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.NotNull(doc);
    }

    [Fact]
    public void CreateFromTemplate_ShouldCreateDocumentFromTemplate()
    {
        var templatePath = CreateWordDocument("test_template.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello <<[ds.Name]>>");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_create_from_template_output.docx");
        var data = new JsonObject { ["Name"] = "World" };
        _tool.Execute("create_from_template", templatePath: templatePath, outputPath: outputPath,
            dataJson: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Hello World", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConvertDocument_ShouldConvertToPdf()
    {
        var docPath = CreateWordDocumentWithContent("test_convert.docx", "Test content");
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        _tool.Execute("convert", path: docPath, outputPath: outputPath, format: "pdf");
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void MergeDocuments_ShouldMergeMultipleDocuments()
    {
        var doc1Path = CreateWordDocumentWithContent("test_merge1.docx", "First document");
        var doc2Path = CreateWordDocumentWithContent("test_merge2.docx", "Second document");
        var outputPath = CreateTestFilePath("test_merge_output.docx");
        _tool.Execute("merge", outputPath: outputPath, inputPaths: [doc1Path, doc2Path]);
        Assert.True(File.Exists(outputPath));
        var mergedDoc = new Document(outputPath);
        var text = mergedDoc.GetText();
        Assert.Contains("First document", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Second document", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SplitDocument_BySection_ShouldSplitBySection()
    {
        var docPath = CreateWordDocument("test_split_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Section 1 content");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Write("Section 2 content");
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("split_section_output");
        Directory.CreateDirectory(outputDir);
        _tool.Execute("split", path: docPath, outputDir: outputDir, splitBy: "section");
        var files = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(files.Length >= 2, $"Expected at least 2 files, got {files.Length}");
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_{operation}.docx");
        var result = _tool.Execute(operation, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("created", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", outputPath: "test.docx"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void ConvertDocument_WithSessionId_ShouldConvert()
    {
        var docPath = CreateWordDocumentWithContent("test_session_convert.docx", "Session content");
        var sessionId = OpenSession(docPath);
        var outputPath = CreateTestFilePath("test_session_convert_output.pdf");
        var result = _tool.Execute("convert", sessionId, outputPath: outputPath, format: "pdf");
        Assert.True(File.Exists(outputPath));
        Assert.Contains("session", result);
    }

    [Fact]
    public void SplitDocument_WithSessionId_ShouldSplit()
    {
        var docPath = CreateWordDocument("test_session_split.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Write("Section 2");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var outputDir = CreateTestFilePath("session_split_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", sessionId, outputDir: outputDir, splitBy: "section");
        var files = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(files.Length >= 2);
        Assert.Contains("session", result);
    }

    [Fact]
    public void CreateFromTemplate_WithSessionId_ShouldCreateFromSessionTemplate()
    {
        var templatePath = CreateWordDocument("test_session_template.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello <<[ds.Name]>>");
        doc.Save(templatePath);

        var sessionId = OpenSession(templatePath);
        var outputPath = CreateTestFilePath("test_session_template_output.docx");
        var data = new JsonObject { ["Name"] = "SessionWorld" };
        var result = _tool.Execute("create_from_template", sessionId, outputPath: outputPath,
            dataJson: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
        Assert.Contains("session", result);
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Hello SessionWorld", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session.pdf");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("convert", "invalid_session_id", outputPath: outputPath, format: "pdf"));
    }

    #endregion
}
