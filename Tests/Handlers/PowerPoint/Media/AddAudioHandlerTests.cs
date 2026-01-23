using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Media;

public class AddAudioHandlerTests : PptHandlerTestBase
{
    private readonly AddAudioHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddAudio()
    {
        Assert.Equal("add_audio", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsAudio()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Audio embedded", result.Message);
        Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("slide 0", result.Message);
    }

    [Fact]
    public void Execute_ReturnsPosition()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile },
            { "x", 200f },
            { "y", 150f }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("200", result.Message);
        Assert.Contains("150", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutAudioPath_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("audioPath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", "/nonexistent/audio.mp3" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile },
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
