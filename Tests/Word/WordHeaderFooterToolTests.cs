using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordHeaderFooterToolTests : WordTestBase
{
    private readonly WordHeaderFooterTool _tool = new();

    [Fact]
    public async Task SetHeaderText_ShouldSetHeaderText()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_header_text.docx");
        var outputPath = CreateTestFilePath("test_set_header_text_output.docx");
        var arguments = CreateArguments("set_header_text", docPath, outputPath);
        arguments["headerLeft"] = "Left Header";
        arguments["headerCenter"] = "Center Header";
        arguments["headerRight"] = "Right Header";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        Assert.Contains("Left", header.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SetFooterText_ShouldSetFooterText()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_footer_text.docx");
        var outputPath = CreateTestFilePath("test_set_footer_text_output.docx");
        var arguments = CreateArguments("set_footer_text", docPath, outputPath);
        arguments["footerLeft"] = "Page";
        arguments["footerRight"] = "{PAGE}";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        Assert.Contains("Page", footer.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetHeadersFooters_ShouldReturnHeadersFooters()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_headers_footers.docx");
        var doc = new Document(docPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
        {
            header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            doc.FirstSection.HeadersFooters.Add(header);
        }

        header.AppendParagraph("Test Header");
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Header", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SetHeaderImage_ShouldSetHeaderImage()
    {
        // Skip test if running in evaluation mode as image operations may be limited
        if (IsEvaluationMode()) return;

        // Arrange
        var docPath = CreateWordDocument("test_set_header_image.docx");

        // Create a simple test image
        var imagePath = CreateTestFilePath("test_header_image.png");
        CreateTestImage(imagePath);

        var outputPath = CreateTestFilePath("test_set_header_image_output.docx");
        var arguments = CreateArguments("set_header_image", docPath, outputPath);
        arguments["imagePath"] = imagePath;
        arguments["alignment"] = "center";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public async Task SetHeaderLine_ShouldSetHeaderLine()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_header_line.docx");
        var outputPath = CreateTestFilePath("test_set_header_line_output.docx");
        var arguments = CreateArguments("set_header_line", docPath, outputPath);
        arguments["lineStyle"] = "single";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
    }

    [Fact]
    public async Task SetFooterLine_ShouldSetFooterLine()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_footer_line.docx");
        var outputPath = CreateTestFilePath("test_set_footer_line_output.docx");
        var arguments = CreateArguments("set_footer_line", docPath, outputPath);
        arguments["lineStyle"] = "single";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
    }

    [Fact]
    public async Task SetHeaderFooter_ShouldSetBoth()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_header_footer.docx");
        var outputPath = CreateTestFilePath("test_set_header_footer_output.docx");
        var arguments = CreateArguments("set_header_footer", docPath, outputPath);
        arguments["headerLeft"] = "Left Header";
        arguments["footerCenter"] = "Center Footer";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(header);
        Assert.NotNull(footer);
        Assert.Contains("Left", header.GetText(), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Center", footer.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Creates a simple test image file for testing
    /// </summary>
    private void CreateTestImage(string imagePath)
    {
        // Create a minimal valid PNG file (1x1 pixel white)
        byte[] pngBytes =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 dimensions
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, // bit depth, color type, etc.
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0xFF, // compressed data
            0x00, 0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59, //
            0xE7, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82
        ];
        File.WriteAllBytes(imagePath, pngBytes);
    }
}