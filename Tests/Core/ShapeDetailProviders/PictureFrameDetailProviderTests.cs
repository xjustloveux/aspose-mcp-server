using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class PictureFrameDetailProviderTests : TestBase
{
    private readonly PictureFrameDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnPicture()
    {
        Assert.Equal("Picture", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }

    [Fact]
    public void CanHandle_WithPictureFrame_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        using var imageStream = new MemoryStream(CreateMinimalPngFile());
        var image = presentation.Images.AddImage(imageStream);
        var pictureFrame = slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 10, 10, 100, 100, image);

        var result = _provider.CanHandle(pictureFrame);

        Assert.True(result);
    }

    [Fact]
    public void GetDetails_WithPictureFrame_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        using var imageStream = new MemoryStream(CreateMinimalPngFile());
        var image = presentation.Images.AddImage(imageStream);
        var pictureFrame = slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 10, 10, 100, 100, image);

        var details = _provider.GetDetails(pictureFrame, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonPictureFrame_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }

    private static byte[] CreateMinimalPngFile()
    {
        using var ms = new MemoryStream();
        // ReSharper disable UseUtf8StringLiteral - Binary PNG data, not text
        ms.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        ms.Write([0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52]);
        ms.Write([0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01]);
        ms.Write([0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE]);
        ms.Write([0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54]);
        ms.Write([0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0x3F, 0x00]);
        ms.Write([0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59, 0xE7]);
        ms.Write([0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44]);
        ms.Write([0xAE, 0x42, 0x60, 0x82]);
        // ReSharper restore UseUtf8StringLiteral
        return ms.ToArray();
    }
}
