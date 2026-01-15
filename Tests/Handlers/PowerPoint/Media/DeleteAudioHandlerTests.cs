using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Media;

public class DeleteAudioHandlerTests : PptHandlerTestBase
{
    private readonly DeleteAudioHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteAudio()
    {
        Assert.Equal("delete_audio", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithAudio(string audioPath)
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        using var audioStream = new FileStream(audioPath, FileMode.Open, FileAccess.Read);
        var audio = pres.Audios.AddAudio(audioStream, LoadingStreamBehavior.ReadStreamAndRelease);
        slide.Shapes.AddAudioFrameEmbedded(100, 100, 50, 50, audio);
        return pres;
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesAudio()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Equal(initialShapeCount - 1, pres.Slides[0].Shapes.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 0", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(99)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", invalidIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonAudioShape_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Test");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not an audio frame", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
