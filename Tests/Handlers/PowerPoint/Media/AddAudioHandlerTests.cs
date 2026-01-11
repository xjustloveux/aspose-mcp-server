using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Audio embedded", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 0", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("200", result);
        Assert.Contains("150", result);
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
