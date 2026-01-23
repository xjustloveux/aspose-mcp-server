using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Results.Word.Image;
using AsposeMcpServer.Tests.Infrastructure;

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
    public void Execute_WithAllSections_ReturnsNullOrOmitsSectionIndex()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesWordResult>(res);

        Assert.Null(result.SectionIndex);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesWordResult>(res);

        Assert.Equal(0, result.Count);
        Assert.NotNull(result.Message);
        Assert.Contains("No images found", result.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsValidStructure()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesWordResult>(res);

        Assert.NotNull(result.Images);
        Assert.True(result.Count >= 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesWordResult>(res);

        Assert.NotNull(result.SectionIndex);
        Assert.Equal(0, result.SectionIndex);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesWordResult>(res);

        Assert.Equal(1, result.Count);
        Assert.Single(result.Images);
        var firstImage = result.Images[0];
        Assert.Equal(0, firstImage.Index);
        Assert.True(firstImage.Width > 0);
        Assert.True(firstImage.Height > 0);
    }

    [Fact]
    public void Execute_ReturnsImageIndex()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesWordResult>(res);

        var firstImage = result.Images[0];
        Assert.Equal(0, firstImage.Index);
    }

    #endregion
}
