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
}