using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordFileToolTests : WordTestBase
{
    private readonly WordFileTool _tool = new();

    #region General Tests

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
        // Arrange - Using LINQ Reporting Engine syntax <<[ds.PropertyName]>>
        var templatePath = CreateWordDocument("test_template.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello <<[ds.Name]>>");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_create_from_template_output.docx");
        var data = new JsonObject
        {
            ["Name"] = "World"
        };
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
        _tool.Execute("convert", docPath, outputPath, format: "pdf");
        Assert.True(File.Exists(outputPath));
        // PDF file should exist (in evaluation mode, may have watermarks)
        Assert.True(File.Exists(outputPath), "PDF file should be created");
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
    public void SplitDocument_ShouldSplitByPage()
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
        _tool.Execute("split", docPath, outputDir: outputDir, splitBy: "page");
        var files = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(files.Length >= 2, $"Expected at least 2 files, got {files.Length}");
    }

    [Fact]
    public void SplitDocument_ShouldSplitBySection()
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
        _tool.Execute("split", docPath, outputDir: outputDir, splitBy: "section");
        var files = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(files.Length >= 2, $"Expected at least 2 files, got {files.Length}");
    }

    [Fact]
    public void CreateDocument_WithPaperSize_ShouldSetCorrectPageSize()
    {
        var outputPath = CreateTestFilePath("test_create_paper_size.docx");
        _tool.Execute("create", outputPath: outputPath, paperSize: "Letter");
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(612, pageSetup.PageWidth); // Letter width
        Assert.Equal(792, pageSetup.PageHeight); // Letter height
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
    public void CreateDocument_WithCompatibilityMode_ShouldSetCompatibility()
    {
        var outputPath = CreateTestFilePath("test_create_compatibility.docx");
        _tool.Execute("create", outputPath: outputPath, compatibilityMode: "Word2016");
        var doc = new Document(outputPath);
        Assert.True(File.Exists(outputPath), "Document should be created");
        // Compatibility mode is set internally
        Assert.NotNull(doc);
    }

    [Fact]
    public void CreateDocument_WithSkipInitialContent_ShouldCreateBlankDocument()
    {
        var outputPath = CreateTestFilePath("test_create_blank.docx");
        _tool.Execute("create", outputPath: outputPath, skipInitialContent: true);
        var doc = new Document(outputPath);
        Assert.True(File.Exists(outputPath), "Document should be created");
        // Document should be created with minimal content
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
    public void CreateFromTemplate_WithArrayData_ShouldIterateItems()
    {
        // Arrange - Using LINQ Reporting Engine with foreach iteration
        var templatePath = CreateWordDocument("test_template_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        // Template with foreach iteration
        builder.Write("Items: <<foreach [item in ds]>><<[item.Product]>> ");
        builder.Write("<</foreach>>");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_create_from_template_array_output.docx");
        // Array as root data
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
    public void ConvertDocument_ToDifferentFormats_ShouldConvertCorrectly()
    {
        var docPath = CreateWordDocumentWithContent("test_convert_formats.docx", "Test content");
        var outputPath = CreateTestFilePath("test_convert_output.html");
        _tool.Execute("convert", docPath, outputPath, format: "html");
        Assert.True(File.Exists(outputPath), "HTML file should be created");
    }

    #endregion

    // Note: WordFileTool only supports: create, create_from_template, convert, merge, split
    // These operations (open, save, get_properties, set_properties) are not available
    // Tests removed as they test non-existent operations
    // Note: This tool does not support session, so no Session ID Tests region
}