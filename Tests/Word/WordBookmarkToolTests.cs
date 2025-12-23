using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordBookmarkToolTests : WordTestBase
{
    private readonly WordBookmarkTool _tool = new();

    [Fact]
    public async Task AddBookmark_ShouldAddBookmark()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_bookmark.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_bookmark_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["name"] = "TestBookmark";
        arguments["text"] = "Bookmarked text";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var bookmark = doc.Range.Bookmarks["TestBookmark"];
        Assert.NotNull(bookmark);
        Assert.Contains("Bookmarked text", bookmark.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetBookmarks_ShouldReturnAllBookmarks()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_bookmarks.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("Bookmark1");
        builder.EndBookmark("Bookmark1");
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

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
        var docPath = CreateWordDocument("test_delete_bookmark.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("BookmarkToDelete");
        builder.EndBookmark("BookmarkToDelete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_bookmark_output.docx");
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["name"] = "BookmarkToDelete";
        arguments["keepText"] = true;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        // Verify bookmark was deleted - try to access it, should throw or return null
        Bookmark? bookmark = null;
        try
        {
            bookmark = resultDoc.Range.Bookmarks["BookmarkToDelete"];
        }
        catch
        {
            // Bookmark not found, which is expected
        }

        Assert.Null(bookmark);
    }

    [Fact]
    public async Task GotoBookmark_ShouldNavigateToBookmark()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_goto_bookmark.docx", "Content before bookmark");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Content after bookmark");
        builder.StartBookmark("TargetBookmark");
        builder.EndBookmark("TargetBookmark");
        doc.Save(docPath);

        var arguments = CreateArguments("goto", docPath);
        arguments["name"] = "TargetBookmark";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("TargetBookmark", result, StringComparison.OrdinalIgnoreCase);
    }
}