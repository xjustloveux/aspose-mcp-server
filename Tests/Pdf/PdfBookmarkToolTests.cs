using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfBookmarkToolTests : PdfTestBase
{
    private readonly PdfBookmarkTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task AddBookmark_ShouldAddBookmark()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_bookmark.pdf");
        var outputPath = CreateTestFilePath("test_add_bookmark_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["title"] = "Chapter 1",
            ["pageIndex"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        var bookmarks = document.Outlines;
        Assert.True(bookmarks.Count > 0, "Document should contain at least one bookmark");
    }

    [Fact]
    public async Task GetBookmarks_ShouldReturnAllBookmarks()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_bookmarks.pdf");
        var document = new Document(pdfPath);
        var outlineItem = new OutlineItemCollection(document.Outlines);
        outlineItem.Title = "Test Bookmark";
        outlineItem.Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1);
        document.Outlines.Add(outlineItem);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Bookmark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteBookmark_ShouldDeleteBookmark()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_bookmark.pdf");
        var document = new Document(pdfPath);
        var outlineItem = new OutlineItemCollection(document.Outlines);
        outlineItem.Title = "Bookmark to Delete";
        outlineItem.Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1);
        document.Outlines.Add(outlineItem);
        document.Save(pdfPath);

        var bookmarksBefore = document.Outlines.Count;
        Assert.True(bookmarksBefore > 0, "Bookmark should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_bookmark_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["bookmarkIndex"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var bookmarksAfter = resultDocument.Outlines.Count;
        Assert.True(bookmarksAfter < bookmarksBefore,
            $"Bookmark should be deleted. Before: {bookmarksBefore}, After: {bookmarksAfter}");
    }
}