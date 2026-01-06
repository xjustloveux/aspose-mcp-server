using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordFileToolTests : WordTestBase
{
    private readonly WordFileTool _tool;

    public WordFileToolTests()
    {
        _tool = new WordFileTool(SessionManager);
    }

    #region General

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
    public void CreateDocument_WithContent_ShouldIncludeContent()
    {
        var outputPath = CreateTestFilePath("test_create_with_content.docx");
        _tool.Execute("create", outputPath: outputPath, content: "Initial content");
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Initial content", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateDocument_WithSkipInitialContent_ShouldCreateBlankDocument()
    {
        var outputPath = CreateTestFilePath("test_create_blank.docx");
        _tool.Execute("create", outputPath: outputPath, skipInitialContent: true);
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.NotNull(doc);
    }

    [Fact]
    public void CreateDocument_WithCustomMargins_ShouldSetMargins()
    {
        var outputPath = CreateTestFilePath("test_create_margins.docx");
        _tool.Execute("create", outputPath: outputPath,
            marginTop: 100.0, marginBottom: 100.0, marginLeft: 80.0, marginRight: 80.0);
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(100, pageSetup.TopMargin);
        Assert.Equal(100, pageSetup.BottomMargin);
        Assert.Equal(80, pageSetup.LeftMargin);
        Assert.Equal(80, pageSetup.RightMargin);
    }

    [Fact]
    public void CreateDocument_WithHeaderFooterDistance_ShouldSetDistances()
    {
        var outputPath = CreateTestFilePath("test_create_header_footer_distance.docx");
        _tool.Execute("create", outputPath: outputPath, headerDistance: 50.0, footerDistance: 50.0);
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(50, pageSetup.HeaderDistance);
        Assert.Equal(50, pageSetup.FooterDistance);
    }

    [Fact]
    public void CreateDocument_WithCustomPageSize_ShouldSetPageSize()
    {
        var outputPath = CreateTestFilePath("test_create_custom_size.docx");
        _tool.Execute("create", outputPath: outputPath, pageWidth: 400.0, pageHeight: 600.0);
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(400, pageSetup.PageWidth);
        Assert.Equal(600, pageSetup.PageHeight);
    }

    [Theory]
    [InlineData("A4", 595.3, 841.9)]
    [InlineData("Letter", 612, 792)]
    [InlineData("A3", 841.9, 1190.55)]
    [InlineData("Legal", 612, 1008)]
    public void CreateDocument_WithDifferentPaperSizes_ShouldSetCorrectSize(string paperSize, double expectedWidth,
        double expectedHeight)
    {
        var outputPath = CreateTestFilePath($"test_create_{paperSize}.docx");
        _tool.Execute("create", outputPath: outputPath, paperSize: paperSize);
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(expectedWidth, pageSetup.PageWidth, 1);
        Assert.Equal(expectedHeight, pageSetup.PageHeight, 1);
    }

    [Theory]
    [InlineData("Word2019")]
    [InlineData("Word2016")]
    [InlineData("Word2013")]
    [InlineData("Word2010")]
    [InlineData("Word2007")]
    public void CreateDocument_WithDifferentCompatibilityModes_ShouldWork(string mode)
    {
        var outputPath = CreateTestFilePath($"test_create_{mode}.docx");
        _tool.Execute("create", outputPath: outputPath, compatibilityMode: mode);
        Assert.True(File.Exists(outputPath));
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
    public void CreateFromTemplate_WithArrayData_ShouldIterateItems()
    {
        var templatePath = CreateWordDocument("test_template_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Items: <<foreach [item in ds]>><<[item.Product]>> <</foreach>>");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_create_from_template_array_output.docx");
        var data = new JsonArray
        {
            new JsonObject { ["Product"] = "Apple" },
            new JsonObject { ["Product"] = "Banana" },
            new JsonObject { ["Product"] = "Cherry" }
        };
        _tool.Execute("create_from_template", templatePath: templatePath, outputPath: outputPath,
            dataJson: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Apple", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Banana", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Cherry", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConvertDocument_ShouldConvertToPdf()
    {
        var docPath = CreateWordDocumentWithContent("test_convert.docx", "Test content");
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        _tool.Execute("convert", path: docPath, outputPath: outputPath, format: "pdf");
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("pdf")]
    [InlineData("html")]
    [InlineData("txt")]
    [InlineData("rtf")]
    public void ConvertDocument_ToDifferentFormats_ShouldWork(string format)
    {
        var docPath = CreateWordDocumentWithContent($"test_convert_{format}.docx", "Test content");
        var outputPath = CreateTestFilePath($"test_convert_output.{format}");
        var result = _tool.Execute("convert", path: docPath, outputPath: outputPath, format: format);
        Assert.True(File.Exists(outputPath));
        Assert.Contains(format, result);
    }

    [Fact]
    public void ConvertDocument_WithoutFormat_ShouldInferFromExtension()
    {
        var docPath = CreateWordDocumentWithContent("test_convert_infer.docx", "Test content");
        var outputPath = CreateTestFilePath("test_convert_infer_output.html");
        var result = _tool.Execute("convert", path: docPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("html", result);
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

    [Theory]
    [InlineData("KeepSourceFormatting")]
    [InlineData("UseDestinationStyles")]
    [InlineData("KeepDifferentStyles")]
    public void MergeDocuments_WithDifferentImportModes_ShouldWork(string mode)
    {
        var doc1Path = CreateWordDocumentWithContent($"test_merge1_{mode}.docx", "First");
        var doc2Path = CreateWordDocumentWithContent($"test_merge2_{mode}.docx", "Second");
        var outputPath = CreateTestFilePath($"test_merge_{mode}_output.docx");
        var result = _tool.Execute("merge", outputPath: outputPath, inputPaths: [doc1Path, doc2Path],
            importFormatMode: mode);
        Assert.True(File.Exists(outputPath));
        Assert.Contains(mode, result);
    }

    [Fact]
    public void MergeDocuments_WithUnlinkHeadersFooters_ShouldUnlink()
    {
        var doc1Path = CreateWordDocumentWithContent("test_merge_unlink1.docx", "First");
        var doc2Path = CreateWordDocumentWithContent("test_merge_unlink2.docx", "Second");
        var outputPath = CreateTestFilePath("test_merge_unlink_output.docx");
        _tool.Execute("merge", outputPath: outputPath, inputPaths: [doc1Path, doc2Path],
            unlinkHeadersFooters: true);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SplitDocument_ByPage_ShouldSplitByPage()
    {
        var docPath = CreateWordDocument("test_split.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 2 content");
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("split_output");
        Directory.CreateDirectory(outputDir);
        _tool.Execute("split", path: docPath, outputDir: outputDir, splitBy: "page");
        var files = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(files.Length >= 2, $"Expected at least 2 files, got {files.Length}");
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

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive_Create(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_{operation}.docx");
        var result = _tool.Execute(operation, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("created", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("CONVERT")]
    [InlineData("Convert")]
    [InlineData("convert")]
    public void Operation_ShouldBeCaseInsensitive_Convert(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, path: docPath, outputPath: outputPath, format: "pdf");
        Assert.True(File.Exists(outputPath));
        Assert.Contains("converted", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("MERGE")]
    [InlineData("Merge")]
    [InlineData("merge")]
    public void Operation_ShouldBeCaseInsensitive_Merge(string operation)
    {
        var doc1Path = CreateWordDocumentWithContent($"test_case_{operation}_1.docx", "First");
        var doc2Path = CreateWordDocumentWithContent($"test_case_{operation}_2.docx", "Second");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, outputPath: outputPath, inputPaths: [doc1Path, doc2Path]);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Merged", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", outputPath: "test.docx"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void CreateDocument_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("create"));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void CreateFromTemplate_WithoutTemplatePath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_template_missing.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("create_from_template", outputPath: outputPath, dataJson: "{}"));
        Assert.Contains("templatePath or sessionId is required", ex.Message);
    }

    [Fact]
    public void CreateFromTemplate_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var templatePath = CreateWordDocument("test_template.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("create_from_template", templatePath: templatePath, dataJson: "{}"));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void CreateFromTemplate_WithoutDataJson_ShouldThrowArgumentException()
    {
        var templatePath = CreateWordDocument("test_template.docx");
        var outputPath = CreateTestFilePath("test_template_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("create_from_template", templatePath: templatePath, outputPath: outputPath));
        Assert.Contains("dataJson parameter is required", ex.Message);
    }

    [Fact]
    public void CreateFromTemplate_WithNonExistentTemplate_ShouldThrowFileNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_template_output.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("create_from_template", templatePath: "nonexistent.docx",
                outputPath: outputPath, dataJson: "{}"));
    }

    [Fact]
    public void ConvertDocument_WithoutPath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", outputPath: outputPath, format: "pdf"));
        Assert.Contains("path or sessionId is required", ex.Message);
    }

    [Fact]
    public void ConvertDocument_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_convert.docx", "Test");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", path: docPath, format: "pdf"));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void ConvertDocument_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_convert_unsupported.docx", "Test");
        var outputPath = CreateTestFilePath("test_convert_unsupported.xyz");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", path: docPath, outputPath: outputPath, format: "xyz"));
        Assert.Contains("Unsupported format", ex.Message);
    }

    [Fact]
    public void MergeDocuments_WithoutInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", outputPath: outputPath));
        Assert.Contains("inputPaths is required", ex.Message);
    }

    [Fact]
    public void MergeDocuments_WithEmptyInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", outputPath: outputPath, inputPaths: []));
        Assert.Contains("inputPaths is required", ex.Message);
    }

    [Fact]
    public void MergeDocuments_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var doc1Path = CreateWordDocumentWithContent("test_merge1.docx", "First");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", inputPaths: [doc1Path]));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void SplitDocument_WithoutPath_ShouldThrowArgumentException()
    {
        var outputDir = CreateTestFilePath("split_output");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", outputDir: outputDir));
        Assert.Contains("path or sessionId is required", ex.Message);
    }

    [Fact]
    public void SplitDocument_WithoutOutputDir_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_split.docx", "Test");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", path: docPath));
        Assert.Contains("outputDir is required", ex.Message);
    }

    #endregion

    #region Session

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
    public void ConvertDocument_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session.pdf");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("convert", "invalid_session_id", outputPath: outputPath, format: "pdf"));
    }

    [Fact]
    public void SplitDocument_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputDir = CreateTestFilePath("invalid_session_split");
        Directory.CreateDirectory(outputDir);
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("split", "invalid_session_id", outputDir: outputDir));
    }

    [Fact]
    public void CreateFromTemplate_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session_template.docx");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("create_from_template", "invalid_session_id",
                outputPath: outputPath, dataJson: "{}"));
    }

    #endregion
}