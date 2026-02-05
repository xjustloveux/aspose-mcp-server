using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Base class for PowerPoint Handler tests providing PowerPoint-specific test infrastructure.
/// </summary>
public abstract class PptHandlerTestBase : HandlerTestBase<Presentation>
{
    /// <summary>
    ///     Creates a new empty PowerPoint presentation for testing.
    /// </summary>
    /// <returns>A new empty Presentation instance.</returns>
    protected static Presentation CreateEmptyPresentation()
    {
        return new Presentation();
    }

    /// <summary>
    ///     Creates a PowerPoint presentation with the specified number of slides.
    /// </summary>
    /// <param name="slideCount">The number of slides to create.</param>
    /// <returns>A Presentation with the specified slides.</returns>
    protected static Presentation CreatePresentationWithSlides(int slideCount)
    {
        var pres = new Presentation();
        for (var i = 1; i < slideCount; i++)
            pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);
        return pres;
    }

    /// <summary>
    ///     Creates a PowerPoint presentation with a text shape.
    ///     The presentation is saved and reloaded via MemoryStream to materialize font metadata,
    ///     ensuring FontsManager.GetFonts() returns the actual fonts used.
    /// </summary>
    /// <param name="text">The text content.</param>
    /// <returns>A Presentation with a text shape on the first slide.</returns>
    protected static Presentation CreatePresentationWithText(string text)
    {
        using var tempPres = new Presentation();
        var slide = tempPres.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 400, 100);
        shape.TextFrame.Text = text;

        using var ms = new MemoryStream();
        tempPres.Save(ms, SaveFormat.Pptx);
        return new Presentation(new MemoryStream(ms.ToArray()));
    }

    /// <summary>
    ///     Asserts that the presentation has the expected number of slides.
    /// </summary>
    /// <param name="pres">The presentation.</param>
    /// <param name="expectedCount">The expected slide count.</param>
    protected static void AssertSlideCount(Presentation pres, int expectedCount)
    {
        Assert.Equal(expectedCount, pres.Slides.Count);
    }

    /// <summary>
    ///     Gets a shape from the first slide.
    /// </summary>
    /// <param name="pres">The presentation.</param>
    /// <param name="shapeIndex">The shape index.</param>
    /// <returns>The shape at the specified index.</returns>
    protected static IShape GetShape(Presentation pres, int shapeIndex)
    {
        return pres.Slides[0].Shapes[shapeIndex];
    }

    /// <summary>
    ///     Creates an operation context with a source path for testing file operations.
    /// </summary>
    /// <param name="pres">The presentation instance.</param>
    /// <param name="sourcePath">The source file path.</param>
    /// <returns>The operation context with source path.</returns>
    protected static OperationContext<Presentation> CreateContextWithPath(Presentation pres, string sourcePath)
    {
        return new OperationContext<Presentation>
        {
            Document = pres,
            SourcePath = sourcePath
        };
    }

    /// <summary>
    ///     Creates a temporary audio file (WAV) for testing.
    /// </summary>
    /// <returns>The full path to the created audio file.</returns>
    protected string CreateTempAudioFile()
    {
        var tempPath = Path.Combine(TestDir, $"test_audio_{Guid.NewGuid()}.wav");
        using var fs = new FileStream(tempPath, FileMode.Create);
        using var bw = new BinaryWriter(fs);
        bw.Write("RIFF".ToCharArray());
        bw.Write(36);
        bw.Write("WAVE".ToCharArray());
        bw.Write("fmt ".ToCharArray());
        bw.Write(16);
        bw.Write((short)1);
        bw.Write((short)1);
        bw.Write(44100);
        bw.Write(88200);
        bw.Write((short)2);
        bw.Write((short)16);
        bw.Write("data".ToCharArray());
        bw.Write(0);
        return tempPath;
    }

    /// <summary>
    ///     Creates a temporary video file (AVI) for testing.
    /// </summary>
    /// <returns>The full path to the created video file.</returns>
    protected string CreateTempVideoFile()
    {
        var tempPath = Path.Combine(TestDir, $"test_video_{Guid.NewGuid()}.avi");
        using var fs = new FileStream(tempPath, FileMode.Create);
        using var bw = new BinaryWriter(fs);
        bw.Write("RIFF".ToCharArray());
        bw.Write(32);
        bw.Write("AVI ".ToCharArray());
        bw.Write("LIST".ToCharArray());
        bw.Write(16);
        bw.Write("hdrl".ToCharArray());
        bw.Write("avih".ToCharArray());
        bw.Write(0);
        return tempPath;
    }
}
