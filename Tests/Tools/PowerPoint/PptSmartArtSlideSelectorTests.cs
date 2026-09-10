using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Covers the SmartArt half of NEW-CONTRACT-02. Both SmartArt handlers read
///     <c>slideIndex</c> with <c>GetRequired</c> and document it as required, but the tool declared
///     it as <c>int slideIndex = 0</c> and forwarded it on every call, so the required check could
///     never fire and omitting the slide silently added or edited SmartArt on the first slide.
/// </summary>
public class PptSmartArtSlideSelectorTests : TestBase
{
    /// <summary>Creates a presentation with the requested number of slides.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="slideCount">Total slides to create.</param>
    /// <returns>The path written.</returns>
    private string CreatePresentation(string fileName, int slideCount)
    {
        var path = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        while (presentation.Slides.Count < slideCount)
            presentation.Slides.AddEmptySlide(presentation.Slides[0].LayoutSlide);
        presentation.Save(path, SaveFormat.Pptx);
        return path;
    }

    /// <summary>Returns the SmartArt shapes on a slide, ignoring any evaluation watermark.</summary>
    /// <param name="slide">The slide to inspect.</param>
    /// <returns>The SmartArt shapes present.</returns>
    private static IReadOnlyList<ISmartArt> SmartArtOn(ISlide slide)
    {
        return slide.Shapes.OfType<ISmartArt>().ToList();
    }

    [Fact]
    public void Add_WithoutSlideIndex_ShouldBeRefused()
    {
        var tool = new PptSmartArtTool();
        var path = CreatePresentation("smartart_missing.pptx", 3);

        Assert.ThrowsAny<ArgumentException>(() =>
            tool.Execute("add", path, layout: "BasicProcess"));

        using var after = new Presentation(path);
        Assert.All(after.Slides, slide => Assert.Empty(SmartArtOn(slide)));
    }

    [Fact]
    public void Add_WithSlideIndex_ShouldStillWork()
    {
        var tool = new PptSmartArtTool();
        var path = CreatePresentation("smartart_ok.pptx", 3);
        var outputPath = CreateTestFilePath("smartart_ok_out.pptx");

        tool.Execute("add", path, outputPath: outputPath, slideIndex: 1, layout: "BasicProcess");

        using var after = new Presentation(outputPath);
        Assert.Empty(SmartArtOn(after.Slides[0]));
        Assert.NotEmpty(SmartArtOn(after.Slides[1]));
    }

    [Fact]
    public void Manage_WithoutSlideIndex_ShouldBeRefused()
    {
        var tool = new PptSmartArtTool();
        var path = CreatePresentation("smartart_manage.pptx", 2);

        Assert.ThrowsAny<ArgumentException>(() =>
            tool.Execute("manage", path, shapeIndex: 0, action: "add", targetPath: "[0]", text: "node"));
    }
}
