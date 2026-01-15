using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Hyperlink;

public class WordHyperlinkHelperTests : WordTestBase
{
    #region GetAllHyperlinks Tests

    [Fact]
    public void GetAllHyperlinks_WithNoHyperlinks_ReturnsEmptyList()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Plain text without hyperlinks");

        var result = WordHyperlinkHelper.GetAllHyperlinks(doc);

        Assert.Empty(result);
    }

    [Fact]
    public void GetAllHyperlinks_WithHyperlink_ReturnsHyperlink()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Click here", "https://example.com", false);

        var result = WordHyperlinkHelper.GetAllHyperlinks(doc);

        Assert.Single(result);
    }

    [Fact]
    public void GetAllHyperlinks_WithMultipleHyperlinks_ReturnsAll()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link 1", "https://example1.com", false);
        builder.Write(" ");
        builder.InsertHyperlink("Link 2", "https://example2.com", false);
        builder.Write(" ");
        builder.InsertHyperlink("Link 3", "https://example3.com", false);

        var result = WordHyperlinkHelper.GetAllHyperlinks(doc);

        Assert.Equal(3, result.Count);
    }

    #endregion

    #region ValidateUrlFormat Tests

    [Theory]
    [InlineData("https://example.com")]
    [InlineData("http://example.com")]
    [InlineData("mailto:test@example.com")]
    [InlineData("ftp://ftp.example.com")]
    [InlineData("file://path/to/file")]
    [InlineData("#bookmark")]
    public void ValidateUrlFormat_WithValidUrl_DoesNotThrow(string url)
    {
        var exception = Record.Exception(() => WordHyperlinkHelper.ValidateUrlFormat(url));

        Assert.Null(exception);
    }

    [Theory]
    [InlineData("HTTPS://EXAMPLE.COM")]
    [InlineData("HTTP://EXAMPLE.COM")]
    [InlineData("MAILTO:TEST@EXAMPLE.COM")]
    public void ValidateUrlFormat_WithUppercasePrefix_DoesNotThrow(string url)
    {
        var exception = Record.Exception(() => WordHyperlinkHelper.ValidateUrlFormat(url));

        Assert.Null(exception);
    }

    [Theory]
    [InlineData("www.example.com")]
    [InlineData("example.com")]
    [InlineData("javascript:alert(1)")]
    [InlineData("invalid-url")]
    public void ValidateUrlFormat_WithInvalidUrl_ThrowsArgumentException(string url)
    {
        var ex = Assert.Throws<ArgumentException>(() => WordHyperlinkHelper.ValidateUrlFormat(url));

        Assert.Contains("Invalid URL format", ex.Message);
    }

    #endregion
}
