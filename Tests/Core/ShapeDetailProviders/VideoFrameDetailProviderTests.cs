using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class VideoFrameDetailProviderTests : TestBase
{
    private readonly VideoFrameDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnVideo()
    {
        Assert.Equal("Video", _provider.TypeName);
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
    public void CanHandle_WithVideoFrame_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        using var videoStream = new MemoryStream(CreateMinimalMp4File());
        var video = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var videoFrame = slide.Shapes.AddVideoFrame(10, 10, 100, 100, video);

        var result = _provider.CanHandle(videoFrame);

        Assert.True(result);
    }

    [Fact]
    public void GetDetails_WithVideoFrame_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        using var videoStream = new MemoryStream(CreateMinimalMp4File());
        var video = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var videoFrame = slide.Shapes.AddVideoFrame(10, 10, 100, 100, video);

        var details = _provider.GetDetails(videoFrame, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonVideoFrame_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }

    private static byte[] CreateMinimalMp4File()
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);
        writer.Write(new byte[] { 0x00, 0x00, 0x00, 0x1C });
        writer.Write("ftyp"u8);
        writer.Write("isom"u8);
        writer.Write(new byte[] { 0x00, 0x00, 0x02, 0x00 });
        writer.Write("isomiso2mp41"u8);
        writer.Write(new byte[] { 0x00, 0x00, 0x00, 0x08 });
        writer.Write("mdat"u8);
        return ms.ToArray();
    }
}
