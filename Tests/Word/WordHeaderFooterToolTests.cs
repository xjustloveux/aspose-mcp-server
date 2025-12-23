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
}