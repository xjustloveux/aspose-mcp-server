using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfRedactToolTests : PdfTestBase
{
    private readonly PdfRedactTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Text to redact"));
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task RedactArea_ShouldRedactArea()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_area.pdf");
        var outputPath = CreateTestFilePath("test_redact_area_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task RedactWithColor_ShouldRedactWithColor()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_color.pdf");
        var outputPath = CreateTestFilePath("test_redact_color_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50,
            ["fillColor"] = "255,0,0"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task RedactWithOverlayText_ShouldRedactWithText()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_overlay.pdf");
        var outputPath = CreateTestFilePath("test_redact_overlay_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50,
            ["overlayText"] = "[REDACTED]"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public async Task RedactWithColorAndOverlayText_ShouldApplyBoth()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_both.pdf");
        var outputPath = CreateTestFilePath("test_redact_both_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50,
            ["fillColor"] = "Red",
            ["overlayText"] = "CONFIDENTIAL"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public async Task Redact_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["pageIndex"] = 99,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task Redact_WithColorName_ShouldParseColor()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_color_name.pdf");
        var outputPath = CreateTestFilePath("test_redact_color_name_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50,
            ["fillColor"] = "Blue"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public async Task Redact_WithDefaultOutput_ShouldOverwriteInput()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_default_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(pdfPath), "PDF should still exist");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public async Task RedactByText_ShouldFindAndRedactText()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_by_text.pdf");
        var outputPath = CreateTestFilePath("test_redact_by_text_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["textToRedact"] = "redact"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redacted", result);
        Assert.Contains("occurrence", result);
    }

    [Fact]
    public async Task RedactByText_OnSpecificPage_ShouldRedactOnlyThatPage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_text_page.pdf");
        var outputPath = CreateTestFilePath("test_redact_text_page_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["textToRedact"] = "Text"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task RedactByText_CaseInsensitive_ShouldFindText()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_case_insensitive.pdf");
        var outputPath = CreateTestFilePath("test_redact_case_insensitive_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["textToRedact"] = "TEXT",
            ["caseSensitive"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task RedactByText_NotFound_ShouldReturnNoOccurrences()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_not_found.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["textToRedact"] = "nonexistent_text_12345"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("No occurrences", result);
        Assert.Contains("No redactions applied", result);
    }

    [Fact]
    public async Task RedactByText_WithOverlayText_ShouldApplyOverlay()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_text_overlay.pdf");
        var outputPath = CreateTestFilePath("test_redact_text_overlay_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["textToRedact"] = "Text",
            ["overlayText"] = "[CLASSIFIED]"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task RedactByText_WithColor_ShouldApplyColor()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_text_color.pdf");
        var outputPath = CreateTestFilePath("test_redact_text_color_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["textToRedact"] = "redact",
            ["fillColor"] = "Red"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task RedactByText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_text_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["pageIndex"] = 99,
            ["textToRedact"] = "Text"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    private string CreateMultiPagePdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        for (var i = 1; i <= 3; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i} contains secret information"));
        }

        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task RedactByText_MultiplePages_ShouldRedactAllPages()
    {
        // Arrange
        var pdfPath = CreateMultiPagePdf("test_redact_multi_page.pdf");
        var outputPath = CreateTestFilePath("test_redact_multi_page_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["textToRedact"] = "secret"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("3 occurrence", result);
        Assert.Contains("3 pages", result);
    }
}