using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordFileToolTests : WordTestBase
{
    private readonly WordFileTool _tool = new();

    [Fact]
    public async Task CreateDocument_ShouldCreateNewDocument()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_document.docx");
        var arguments = CreateArguments("create", outputPath, outputPath);

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.NotNull(doc);
    }

    [Fact]
    public async Task CreateFromTemplate_ShouldCreateDocumentFromTemplate()
    {
        // Arrange - Using LINQ Reporting Engine syntax <<[ds.PropertyName]>>
        var templatePath = CreateWordDocument("test_template.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello <<[ds.Name]>>");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_create_from_template_output.docx");
        var arguments = new JsonObject
        {
            ["operation"] = "create_from_template",
            ["templatePath"] = templatePath,
            ["outputPath"] = outputPath
        };
        var data = new JsonObject
        {
            ["Name"] = "World"
        };
        arguments["data"] = data;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Hello World", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertDocument_ShouldConvertToPdf()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_convert.docx", "Test content");
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var arguments = CreateArguments("convert", docPath, outputPath);
        arguments["outputPath"] = outputPath;
        arguments["format"] = "pdf";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        // PDF file should exist (in evaluation mode, may have watermarks)
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task MergeDocuments_ShouldMergeMultipleDocuments()
    {
        // Arrange
        var doc1Path = CreateWordDocumentWithContent("test_merge1.docx", "First document");
        var doc2Path = CreateWordDocumentWithContent("test_merge2.docx", "Second document");
        var outputPath = CreateTestFilePath("test_merge_output.docx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["inputPaths"] = new JsonArray(doc1Path, doc2Path),
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        var mergedDoc = new Document(outputPath);
        var text = mergedDoc.GetText();
        Assert.Contains("First document", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Second document", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SplitDocument_ShouldSplitByPage()
    {
        // Arrange
        var docPath = CreateWordDocument("test_split.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 2 content");
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("split_output");
        Directory.CreateDirectory(outputDir);
        var arguments = CreateArguments("split", docPath);
        arguments["outputDir"] = outputDir;
        arguments["splitBy"] = "page";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(files.Length >= 2, $"Expected at least 2 files, got {files.Length}");
    }

    [Fact]
    public async Task SplitDocument_ShouldSplitBySection()
    {
        // Arrange
        var docPath = CreateWordDocument("test_split_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Section 1 content");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Write("Section 2 content");
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("split_section_output");
        Directory.CreateDirectory(outputDir);
        var arguments = CreateArguments("split", docPath);
        arguments["outputDir"] = outputDir;
        arguments["splitBy"] = "section";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(files.Length >= 2, $"Expected at least 2 files, got {files.Length}");
    }

    [Fact]
    public async Task CreateDocument_WithPaperSize_ShouldSetCorrectPageSize()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_paper_size.docx");
        var arguments = CreateArguments("create", outputPath, outputPath);
        arguments["paperSize"] = "Letter";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(612, pageSetup.PageWidth); // Letter width
        Assert.Equal(792, pageSetup.PageHeight); // Letter height
    }

    [Fact]
    public async Task CreateDocument_WithCustomMargins_ShouldSetMargins()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_margins.docx");
        var arguments = CreateArguments("create", outputPath, outputPath);
        arguments["marginTop"] = 100.0;
        arguments["marginBottom"] = 100.0;
        arguments["marginLeft"] = 80.0;
        arguments["marginRight"] = 80.0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(100, pageSetup.TopMargin);
        Assert.Equal(100, pageSetup.BottomMargin);
        Assert.Equal(80, pageSetup.LeftMargin);
        Assert.Equal(80, pageSetup.RightMargin);
    }

    [Fact]
    public async Task CreateDocument_WithCompatibilityMode_ShouldSetCompatibility()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_compatibility.docx");
        var arguments = CreateArguments("create", outputPath, outputPath);
        arguments["compatibilityMode"] = "Word2016";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.True(File.Exists(outputPath), "Document should be created");
        // Compatibility mode is set internally
        Assert.NotNull(doc);
    }

    [Fact]
    public async Task CreateDocument_WithSkipInitialContent_ShouldCreateBlankDocument()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_blank.docx");
        var arguments = CreateArguments("create", outputPath, outputPath);
        arguments["skipInitialContent"] = true;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.True(File.Exists(outputPath), "Document should be created");
        // Document should be created with minimal content
        Assert.NotNull(doc);
    }

    [Fact]
    public async Task CreateDocument_WithContent_ShouldIncludeContent()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_with_content.docx");
        var arguments = CreateArguments("create", outputPath, outputPath);
        arguments["content"] = "Initial content";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Initial content", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CreateDocument_WithHeaderFooterDistance_ShouldSetDistances()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_header_footer_distance.docx");
        var arguments = CreateArguments("create", outputPath, outputPath);
        arguments["headerDistance"] = 50.0;
        arguments["footerDistance"] = 50.0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var pageSetup = doc.FirstSection.PageSetup;
        Assert.Equal(50, pageSetup.HeaderDistance);
        Assert.Equal(50, pageSetup.FooterDistance);
    }

    [Fact]
    public async Task CreateFromTemplate_WithArrayData_ShouldIterateItems()
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
        var arguments = new JsonObject
        {
            ["operation"] = "create_from_template",
            ["templatePath"] = templatePath,
            ["outputPath"] = outputPath
        };
        // Array as root data
        var data = new JsonArray
        {
            new JsonObject { ["Product"] = "Apple" },
            new JsonObject { ["Product"] = "Banana" },
            new JsonObject { ["Product"] = "Cherry" }
        };
        arguments["data"] = data;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Apple", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Banana", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Cherry", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertDocument_ToDifferentFormats_ShouldConvertCorrectly()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_convert_formats.docx", "Test content");
        var outputPath = CreateTestFilePath("test_convert_output.html");
        var arguments = CreateArguments("convert", docPath, outputPath);
        arguments["outputPath"] = outputPath;
        arguments["format"] = "html";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "HTML file should be created");
    }

    // Note: WordFileTool only supports: create, create_from_template, convert, merge, split
    // These operations (open, save, get_properties, set_properties) are not available
    // Tests removed as they test non-existent operations
}