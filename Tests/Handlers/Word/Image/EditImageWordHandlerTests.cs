using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Tests.Handlers.Word.Image;

public class EditImageWordHandlerTests : WordHandlerTestBase
{
    private readonly EditImageWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsImageSize()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "width", 200.0 },
            { "height", 150.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("width", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("height", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsAlternativeText()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "alternativeText", "New alt text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("alt text", result.Message, StringComparison.OrdinalIgnoreCase);
        var shape = GetFirstImage(doc);
        Assert.Equal("New alt text", shape.AlternativeText);
    }

    [Fact]
    public void Execute_EditsTitle()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "title", "New title" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("title", result.Message, StringComparison.OrdinalIgnoreCase);
        var shape = GetFirstImage(doc);
        Assert.Equal("New title", shape.Title);
    }

    [Fact]
    public void Execute_EditsLinkUrl()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "linkUrl", "https://example.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("hyperlink", result.Message, StringComparison.OrdinalIgnoreCase);
        var shape = GetFirstImage(doc);
        Assert.Equal("https://example.com", shape.HRef);
    }

    [Fact]
    public void Execute_WithAspectRatioLocked_SetsAspectRatio()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "aspectRatioLocked", true }
        });

        _handler.Execute(context, parameters);

        var shape = GetFirstImage(doc);
        Assert.True(shape.AspectRatioLocked);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidImageIndex_ThrowsArgumentException()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeImageIndex_ThrowsArgumentException()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithImage(string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        return doc;
    }

    private static WordShape GetFirstImage(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<WordShape>()
            .First(s => s.HasImage);
    }

    #endregion
}
