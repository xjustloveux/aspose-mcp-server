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
        using var document = new Document();
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Added bookmark", result);
        using var document = new Document(outputPath);
        Assert.True(document.Outlines.Count > 0, "Document should contain at least one bookmark");
    }

    [Fact]
    public async Task AddBookmark_InvalidPageIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["title"] = "Test",
            ["pageIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetBookmarks_WithBookmarks_ShouldReturnBookmarkInfo()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_bookmarks.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Test Bookmark",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 1", result);
        Assert.Contains("Test Bookmark", result);
    }

    [Fact]
    public async Task GetBookmarks_Empty_ShouldReturnEmptyResult()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No bookmarks found", result);
    }

    [Fact]
    public async Task DeleteBookmark_ShouldDeleteBookmark()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_bookmark.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Bookmark to Delete",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_delete_bookmark_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["bookmarkIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Deleted bookmark", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.Outlines);
    }

    [Fact]
    public async Task DeleteBookmark_InvalidIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_invalid.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["bookmarkIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task EditBookmark_ShouldEditTitle()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_bookmark.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Original Title",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_bookmark_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["bookmarkIndex"] = 1,
            ["title"] = "Updated Title"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Edited bookmark", result);
        using var resultDocument = new Document(outputPath);
        Assert.Equal("Updated Title", resultDocument.Outlines[1].Title);
    }

    [Fact]
    public async Task EditBookmark_ShouldEditPageIndex()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_page.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Bookmark",
                Action = new GoToAction(document.Pages[1])
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["bookmarkIndex"] = 1,
            ["pageIndex"] = 2
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Edited bookmark", result);
    }

    [Fact]
    public async Task EditBookmark_InvalidBookmarkIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_invalid.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["bookmarkIndex"] = 99,
            ["title"] = "Test"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task EditBookmark_InvalidPageIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_invalid_page.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Bookmark",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["bookmarkIndex"] = 1,
            ["pageIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}