using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Tests.Helpers;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Tests.Handlers.Word.Image;

public class AddImageWordHandlerTests : WordHandlerTestBase
{
    private readonly AddImageWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static int GetImageCount(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<WordShape>()
            .Count(s => s.HasImage);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsImage()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Image added", result);
        Assert.Equal(1, GetImageCount(doc));
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsImageFileName()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains(Path.GetFileName(tempFile), result);
    }

    #endregion

    #region Image Options

    [Fact]
    public void Execute_WithWidthAndHeight_SetsSize()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "width", 150.0 },
            { "height", 100.0 }
        });

        var result = _handler.Execute(context, parameters);

        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();
        Assert.Single(shapes);
        Assert.Contains("Size:", result);
    }

    [Fact]
    public void Execute_WithAlternativeText_SetsAltText()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "alternativeText", "Test image alt text" }
        });

        _handler.Execute(context, parameters);

        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();
        Assert.Single(shapes);
        Assert.Equal("Test image alt text", shapes[0].AlternativeText);
    }

    [Fact]
    public void Execute_WithTitle_SetsTitle()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "title", "Test image title" }
        });

        _handler.Execute(context, parameters);

        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();
        Assert.Single(shapes);
        Assert.Equal("Test image title", shapes[0].Title);
    }

    [Fact]
    public void Execute_WithLinkUrl_SetsHyperlink()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "linkUrl", "https://example.com" }
        });

        _handler.Execute(context, parameters);

        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();
        Assert.Single(shapes);
        Assert.Equal("https://example.com", shapes[0].HRef);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    public void Execute_WithAlignment_ReturnsAlignmentInMessage(string alignment)
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "alignment", alignment }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"Alignment: {alignment}", result);
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("square")]
    [InlineData("tight")]
    public void Execute_WithTextWrapping_ReturnsWrappingInMessage(string textWrapping)
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "textWrapping", textWrapping }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"Text wrapping: {textWrapping}", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutImagePath_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "/nonexistent/path/image.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
