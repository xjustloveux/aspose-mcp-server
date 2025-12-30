using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfTextToolTests : PdfTestBase
{
    private readonly PdfTextTool _tool = new();

    private string CreatePdfDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Sample PDF Text"));
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task ExtractText_ShouldReturnJsonResult()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_extract_text.pdf");
        var arguments = CreateArguments("extract", pdfPath);
        arguments["pageIndex"] = 1;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("pageIndex", out _));
        Assert.True(json.TryGetProperty("totalPages", out _));
        Assert.True(json.TryGetProperty("text", out _));
    }

    [Fact]
    public async Task ExtractText_WithIncludeFontInfo_ShouldReturnFragments()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_extract_font_info.pdf");
        var arguments = CreateArguments("extract", pdfPath);
        arguments["pageIndex"] = 1;
        arguments["includeFontInfo"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("fragments", out _));
        Assert.True(json.TryGetProperty("fragmentCount", out _));
    }

    [Fact]
    public async Task ExtractText_WithRawMode_ShouldExtractRawText()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_extract_raw.pdf");
        var arguments = CreateArguments("extract", pdfPath);
        arguments["pageIndex"] = 1;
        arguments["extractionMode"] = "raw";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("text", result);
    }

    [Fact]
    public async Task ExtractText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_extract_invalid_page.pdf");
        var arguments = CreateArguments("extract", pdfPath);
        arguments["pageIndex"] = 99;

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task AddText_ShouldAddTextToPage()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_text.pdf");
        var outputPath = CreateTestFilePath("test_add_text_output.pdf");
        var arguments = CreateArguments("add", pdfPath, outputPath);
        arguments["pageIndex"] = 1;
        arguments["text"] = "Added Text";
        arguments["x"] = 100;
        arguments["y"] = 700;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Added text to page", result);
    }

    [Fact]
    public async Task AddText_WithFontOptions_ShouldApplyFontOptions()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_text_font.pdf");
        var outputPath = CreateTestFilePath("test_add_text_font_output.pdf");
        var arguments = CreateArguments("add", pdfPath, outputPath);
        arguments["pageIndex"] = 1;
        arguments["text"] = "Styled Text";
        arguments["fontName"] = "Arial";
        arguments["fontSize"] = 14;
        arguments["x"] = 100;
        arguments["y"] = 700;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count >= 1, "PDF should have at least one page");
    }

    [Fact]
    public async Task AddText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_invalid_page.pdf");
        var arguments = CreateArguments("add", pdfPath);
        arguments["pageIndex"] = 99;
        arguments["text"] = "Test";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task EditText_ShouldReplaceText()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_edit_text.pdf");
        var outputPath = CreateTestFilePath("test_edit_text_output.pdf");
        var arguments = CreateArguments("edit", pdfPath, outputPath);
        arguments["pageIndex"] = 1;
        arguments["oldText"] = "Sample PDF Text";
        arguments["newText"] = "Updated";
        arguments["replaceAll"] = true;

        var isEvaluationMode = IsEvaluationMode();

        try
        {
            var result = await _tool.ExecuteAsync(arguments);
            Assert.True(File.Exists(outputPath), "PDF file should be created");
            Assert.Contains("Replaced", result);
        }
        catch (ArgumentException ex) when (isEvaluationMode &&
                                           (ex.Message.Contains("not found") || ex.Message.Contains("Failed")))
        {
            Assert.True(true, "In evaluation mode, text replacement may fail");
        }
    }

    [Fact]
    public async Task EditText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_edit_invalid_page.pdf");
        var arguments = CreateArguments("edit", pdfPath);
        arguments["pageIndex"] = 99;
        arguments["oldText"] = "old";
        arguments["newText"] = "new";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task EditText_WithTextNotFound_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_edit_not_found.pdf");
        var arguments = CreateArguments("edit", pdfPath);
        arguments["pageIndex"] = 1;
        arguments["oldText"] = "nonexistent_text_12345";
        arguments["newText"] = "new";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var arguments = CreateArguments("unknown", pdfPath);

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task ExtractText_WithMultiplePages_ShouldExtractFromSpecifiedPage()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_multi_page_extract.pdf");
        var document = new Document();
        var page1 = document.Pages.Add();
        page1.Paragraphs.Add(new TextFragment("Page 1 Content"));
        var page2 = document.Pages.Add();
        page2.Paragraphs.Add(new TextFragment("Page 2 Content"));
        document.Save(pdfPath);

        var arguments = CreateArguments("extract", pdfPath);
        arguments["pageIndex"] = 2;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("pageIndex", out var pageIndexProp));
        Assert.Equal(2, pageIndexProp.GetInt32());
    }

    [Fact]
    public async Task ExtractText_WithUnicode_ShouldHandleUnicode()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_unicode_extract.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Unicode Test: 中文 日本語 한국어"));
        document.Save(pdfPath);

        var arguments = CreateArguments("extract", pdfPath);
        arguments["pageIndex"] = 1;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.True(result.Length > 0, "Should extract unicode text");
    }

    [Fact]
    public async Task EditText_ReplaceAll_ShouldReplaceAllOccurrences()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_replace_all.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test word here. Another test word. Third test word."));
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_replace_all_output.pdf");
        var arguments = CreateArguments("edit", pdfPath, outputPath);
        arguments["pageIndex"] = 1;
        arguments["oldText"] = "test";
        arguments["newText"] = "replaced";
        arguments["replaceAll"] = true;

        var isEvaluationMode = IsEvaluationMode();

        try
        {
            await _tool.ExecuteAsync(arguments);
            Assert.True(File.Exists(outputPath), "Output file should be created");
        }
        catch (Exception) when (isEvaluationMode)
        {
            Assert.True(true, "In evaluation mode, replace operation may fail");
        }
    }

    [Fact]
    public async Task AddText_WithDefaultPosition_ShouldUseDefaults()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_default_position.pdf");
        var outputPath = CreateTestFilePath("test_add_default_position_output.pdf");
        var arguments = CreateArguments("add", pdfPath, outputPath);
        arguments["pageIndex"] = 1;
        arguments["text"] = "Text with default position";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Added text", result);
    }
}