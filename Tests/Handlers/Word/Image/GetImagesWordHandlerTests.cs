using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Image;

public class GetImagesWordHandlerTests : WordHandlerTestBase
{
    private readonly GetImagesWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Section Index Parameter

    [Fact]
    public void Execute_WithAllSections_ReturnsNullSectionIndex()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("sectionIndex", out var sectionIndexElement));
        Assert.Equal(JsonValueKind.Null, sectionIndexElement.ValueKind);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
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

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoImages_ReturnsEmptyList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.True(json.RootElement.TryGetProperty("message", out var messageElement));
        Assert.Contains("No images found", messageElement.GetString());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsValidJsonStructure()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("images", out _));
    }

    [Fact]
    public void Execute_ReturnsSectionIndex()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("sectionIndex", out var sectionIndexElement));
        Assert.Equal(0, sectionIndexElement.GetInt32());
    }

    #endregion

    #region With Actual Images

    [Fact]
    public void Execute_WithImages_ReturnsImageInfo()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        var images = json.RootElement.GetProperty("images");
        Assert.Equal(1, images.GetArrayLength());
        var firstImage = images[0];
        Assert.True(firstImage.TryGetProperty("index", out _));
        Assert.True(firstImage.TryGetProperty("width", out _));
        Assert.True(firstImage.TryGetProperty("height", out _));
    }

    [Fact]
    public void Execute_ReturnsImageIndex()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var images = json.RootElement.GetProperty("images");
        var firstImage = images[0];
        Assert.Equal(0, firstImage.GetProperty("index").GetInt32());
    }

    #endregion
}
