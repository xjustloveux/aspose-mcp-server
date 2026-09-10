using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.PowerPoint.Image;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;
using SkiaSharp;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Covers TEST-11 for PowerPoint. Image tests worked with a single picture, so nothing proved
///     that the index reported by <c>get</c> addresses the same picture in <c>edit</c> and
///     <c>delete</c>, nor that the remaining pictures keep their order after a deletion. Three
///     frames of distinct sizes make each one identifiable in the results.
/// </summary>
public class PptImageMultiRoundTripTests : TestBase
{
    /// <summary>Writes a PNG of the requested pixel size.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="size">Bitmap width and height in pixels.</param>
    /// <returns>The path written.</returns>
    private string CreatePng(string fileName, int size)
    {
        var path = CreateTestFilePath(fileName);
        using var bitmap = new SKBitmap(size, size);
        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        File.WriteAllBytes(path, data.ToArray());
        return path;
    }

    /// <summary>Builds a presentation whose first slide holds three pictures of widths 40, 80, 120.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <returns>The presentation path.</returns>
    private string CreatePresentationWithThreeImages(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        var y = 10f;
        foreach (var (name, size) in new[] { ("a.png", 40), ("b.png", 80), ("c.png", 120) })
        {
            var image = presentation.Images.AddImage(File.ReadAllBytes(CreatePng(name, size)));
            slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 10, y, size, size, image);
            y += size + 10;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return path;
    }

    /// <summary>Runs the get operation for slide 0 and returns the reported pictures.</summary>
    /// <param name="tool">The tool under test.</param>
    /// <param name="path">Presentation path.</param>
    /// <returns>The pictures reported, in the order the tool reports them.</returns>
    private static IReadOnlyList<PptImageInfo> GetImages(PptImageTool tool, string path)
    {
        var raw = tool.Execute("get", path, slideIndex: 0);
        return ((FinalizedResult<GetImagesPptResult>)raw).Data.Images;
    }

    [Fact]
    public void Get_ShouldReportEveryPictureWithSequentialIndices()
    {
        var tool = new PptImageTool();
        var path = CreatePresentationWithThreeImages("three.pptx");

        var images = GetImages(tool, path);

        Assert.Equal(3, images.Count);
        Assert.Equal([0, 1, 2], images.Select(i => i.ImageIndex));
        Assert.Equal([40f, 80f, 120f], images.Select(i => i.Width));
    }

    [Fact]
    public void Edit_ShouldResizeThePictureThatGetReportedAtThatIndex()
    {
        var tool = new PptImageTool();
        var path = CreatePresentationWithThreeImages("edit.pptx");
        var outputPath = CreateTestFilePath("edit_out.pptx");

        tool.Execute("edit", path, outputPath: outputPath, slideIndex: 0, imageIndex: 1,
            width: 200, height: 200);

        var images = GetImages(tool, outputPath);
        Assert.Equal(3, images.Count);
        Assert.Equal(40f, images[0].Width);
        Assert.Equal(200f, images[1].Width);
        Assert.Equal(120f, images[2].Width);
    }

    [Fact]
    public void Delete_ShouldRemoveTheAddressedPictureAndKeepTheOrderOfTheRest()
    {
        var tool = new PptImageTool();
        var path = CreatePresentationWithThreeImages("delete.pptx");
        var outputPath = CreateTestFilePath("delete_out.pptx");

        tool.Execute("delete", path, outputPath: outputPath, slideIndex: 0, imageIndex: 1);

        var images = GetImages(tool, outputPath);
        Assert.Equal(2, images.Count);
        Assert.Equal([40f, 120f], images.Select(i => i.Width));
        Assert.Equal([0, 1], images.Select(i => i.ImageIndex));
    }

    [Fact]
    public void Delete_ShouldRemoveTheLastPictureWhenAddressedByItsIndex()
    {
        var tool = new PptImageTool();
        var path = CreatePresentationWithThreeImages("delete_last.pptx");
        var outputPath = CreateTestFilePath("delete_last_out.pptx");

        tool.Execute("delete", path, outputPath: outputPath, slideIndex: 0, imageIndex: 2);

        var images = GetImages(tool, outputPath);
        Assert.Equal([40f, 80f], images.Select(i => i.Width));
    }
}
