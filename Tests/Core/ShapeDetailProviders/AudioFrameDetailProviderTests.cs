using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;
using AsposeMcpServer.Core.ShapeDetailProviders.Providers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class AudioFrameDetailProviderTests : TestBase
{
    private readonly AudioFrameDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnAudio()
    {
        Assert.Equal("Audio", _provider.TypeName);
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
    public void CanHandle_WithAudioFrame_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        using var audioStream = new MemoryStream(CreateMinimalWavFile());
        var audio = presentation.Audios.AddAudio(audioStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var audioFrame = slide.Shapes.AddAudioFrameEmbedded(10, 10, 100, 100, audio);

        var result = _provider.CanHandle(audioFrame);

        Assert.True(result);
    }

    [Fact]
    public void GetDetails_WithAudioFrame_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        using var audioStream = new MemoryStream(CreateMinimalWavFile());
        var audio = presentation.Audios.AddAudio(audioStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var audioFrame = slide.Shapes.AddAudioFrameEmbedded(10, 10, 100, 100, audio);

        var details = _provider.GetDetails(audioFrame, presentation);
        var audioDetails = Assert.IsType<AudioFrameDetails>(details);

        Assert.False(string.IsNullOrEmpty(audioDetails.PlayMode));
        Assert.False(string.IsNullOrEmpty(audioDetails.Volume));
    }

    [Fact]
    public void GetDetails_WithNonAudioFrame_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }

    private static byte[] CreateMinimalWavFile()
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);
        writer.Write("RIFF"u8);
        writer.Write(36);
        writer.Write("WAVE"u8);
        writer.Write("fmt "u8);
        writer.Write(16);
        writer.Write((short)1);
        writer.Write((short)1);
        writer.Write(8000);
        writer.Write(8000);
        writer.Write((short)1);
        writer.Write((short)8);
        writer.Write("data"u8);
        writer.Write(0);
        return ms.ToArray();
    }
}
