using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfLinkToolTests : PdfTestBase
{
    private readonly PdfLinkTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task AddLink_ShouldAddLink()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_link.pdf");
        var outputPath = CreateTestFilePath("test_add_link_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 30,
            ["url"] = "https://example.com"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        var page = document.Pages[1];
        var annotations = page.Annotations;
        Assert.True(annotations.Count > 0, "Page should contain at least one link annotation");
    }

    [Fact]
    public async Task GetLinks_ShouldReturnAllLinks()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_links.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://test.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public async Task DeleteLink_ShouldDeleteLink()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_link.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://delete.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var linksBefore = document.Pages[1].Annotations.Count;
        Assert.True(linksBefore > 0, "Link should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_link_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["linkIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var linksAfter = resultDocument.Pages[1].Annotations.Count;
        Assert.True(linksAfter < linksBefore,
            $"Link should be deleted. Before: {linksBefore}, After: {linksAfter}");
    }

    [Fact]
    public async Task EditLink_ShouldEditLink()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_link.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://original.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_edit_link_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["linkIndex"] = 0,
            ["url"] = "https://updated.com"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output file should be created");
        var resultDocument = new Document(outputPath);
        var annotations = resultDocument.Pages[1].Annotations;
        Assert.True(annotations.Count > 0, "Page should still have annotations");
    }

    [Fact]
    public async Task AddLink_WithTargetPage_ShouldAddInternalLink()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_add_internal_link.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_add_internal_link_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 30,
            ["targetPage"] = 2
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var annotations = resultDocument.Pages[1].Annotations.OfType<LinkAnnotation>().ToList();
        Assert.True(annotations.Count > 0, "Page should contain internal link");
        Assert.IsType<GoToAction>(annotations[0].Action);
    }

    [Fact]
    public async Task GetLinks_WithInternalLink_ShouldReturnDestinationPage()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_get_internal_links.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToAction(document.Pages[2])
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"type\": \"page\"", result);
    }

    [Fact]
    public async Task GetLinks_WithNoLinks_ShouldReturnEmptyResult()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_no_links.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No links found", result);
    }

    [Fact]
    public async Task GetLinks_WithoutPageIndex_ShouldReturnAllLinks()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_get_all_links.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var page1 = document.Pages[1];
        var link1 = new LinkAnnotation(page1, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://page1.com")
        };
        page1.Annotations.Add(link1);

        var page2 = document.Pages[2];
        var link2 = new LinkAnnotation(page2, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://page2.com")
        };
        page2.Annotations.Add(link2);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 2", result);
    }

    [Fact]
    public async Task AddLink_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["pageIndex"] = 99,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 30,
            ["url"] = "https://example.com"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task AddLink_WithoutUrlOrTargetPage_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_no_url.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 30
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Either url or targetPage must be provided", exception.Message);
    }

    [Fact]
    public async Task AddLink_WithInvalidTargetPage_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_invalid_target.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 30,
            ["targetPage"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("targetPage must be between", exception.Message);
    }

    [Fact]
    public async Task DeleteLink_WithInvalidLinkIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_invalid_index.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://test.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["linkIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("linkIndex must be between", exception.Message);
    }

    [Fact]
    public async Task EditLink_WithInvalidLinkIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_invalid_index.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["linkIndex"] = 99,
            ["url"] = "https://test.com"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("linkIndex must be between", exception.Message);
    }

    [Fact]
    public async Task EditLink_WithTargetPage_ShouldChangeToInternalLink()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_edit_to_internal.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://original.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_edit_to_internal_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["linkIndex"] = 0,
            ["targetPage"] = 2
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var annotations = resultDocument.Pages[1].Annotations.OfType<LinkAnnotation>().ToList();
        Assert.True(annotations.Count > 0);
        Assert.IsType<GoToAction>(annotations[0].Action);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task GetLinks_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath,
            ["pageIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }
}