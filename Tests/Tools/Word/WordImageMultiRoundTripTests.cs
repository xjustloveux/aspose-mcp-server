using Aspose.Words;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Word.Image;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;
using SkiaSharp;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Covers TEST-11 for Word. Image tests used a single image, so nothing proved that the index
///     reported by <c>get</c> addresses the same image in <c>edit</c> and <c>delete</c>, nor that
///     the survivors keep their order after a deletion. Three images of distinct sizes make each
///     one identifiable in the results.
/// </summary>
public class WordImageMultiRoundTripTests : TestBase
{
    /// <summary>Widths this fixture creates, in points. Anything else came from elsewhere.</summary>
    private static readonly double[] FixtureWidths = [40d, 80d, 120d, 200d, 300d];

    /// <summary>Writes a PNG of the requested pixel size so each image is distinguishable.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="width">Bitmap width in pixels.</param>
    /// <param name="height">Bitmap height in pixels.</param>
    /// <returns>The path written.</returns>
    private string CreatePng(string fileName, int width, int height)
    {
        var path = CreateTestFilePath(fileName);
        using var bitmap = new SKBitmap(width, height);
        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        File.WriteAllBytes(path, data.ToArray());
        return path;
    }

    /// <summary>Builds a document holding three images with widths 40, 80 and 120 points.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <returns>The document path.</returns>
    private string CreateDocumentWithThreeImages(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        var document = new Document();
        var builder = new DocumentBuilder(document);

        foreach (var (name, size) in new[] { ("a.png", 40), ("b.png", 80), ("c.png", 120) })
        {
            var shape = builder.InsertImage(CreatePng(name, size, size));
            shape.Width = size;
            shape.Height = size;
            builder.Writeln();
        }

        document.Save(path, SaveFormat.Docx);
        return path;
    }

    /// <summary>
    ///     Runs the get operation and returns the fixture's own images. An unlicensed build
    ///     appends an evaluation watermark picture after them, which is not part of what these
    ///     tests are about; it is dropped so the same assertions hold in both modes.
    /// </summary>
    /// <param name="tool">The tool under test.</param>
    /// <param name="path">Document path.</param>
    /// <returns>The fixture images, in the order the tool reports them.</returns>
    private static IReadOnlyList<WordImageInfo> GetImages(WordImageTool tool, string path)
    {
        var raw = tool.Execute("get", path);
        return ((FinalizedResult<GetImagesWordResult>)raw).Data.Images
            .Where(image => FixtureWidths.Contains(image.Width)
                            || image.OriginalSize?.WidthPixels == 300)
            .ToList();
    }

    [Fact]
    public void Get_ShouldReportEveryImageWithSequentialIndices()
    {
        var tool = new WordImageTool();
        var path = CreateDocumentWithThreeImages("three.docx");

        var images = GetImages(tool, path);

        Assert.Equal(3, images.Count);
        Assert.Equal([0, 1, 2], images.Select(i => i.Index));
        Assert.Equal([40d, 80d, 120d], images.Select(i => i.Width));
    }

    [Fact]
    public void Edit_ShouldChangeTheImageThatGetReportedAtThatIndex()
    {
        var tool = new WordImageTool();
        var path = CreateDocumentWithThreeImages("edit.docx");
        var outputPath = CreateTestFilePath("edit_out.docx");

        tool.Execute("edit", path, outputPath: outputPath, imageIndex: 1, width: 200, height: 200);

        var images = GetImages(tool, outputPath);
        Assert.Equal(3, images.Count);
        Assert.Equal(40d, images[0].Width);
        Assert.Equal(200d, images[1].Width);
        Assert.Equal(120d, images[2].Width);
    }

    [Fact]
    public void Delete_ShouldRemoveTheAddressedImageAndKeepTheOrderOfTheRest()
    {
        var tool = new WordImageTool();
        var path = CreateDocumentWithThreeImages("delete.docx");
        var outputPath = CreateTestFilePath("delete_out.docx");

        tool.Execute("delete", path, outputPath: outputPath, imageIndex: 1);

        var images = GetImages(tool, outputPath);
        Assert.Equal(2, images.Count);
        Assert.Equal([40d, 120d], images.Select(i => i.Width));
        Assert.Equal([0, 1], images.Select(i => i.Index));
    }

    [Fact]
    public void Delete_ShouldRemoveTheLastImageWhenAddressedByItsIndex()
    {
        var tool = new WordImageTool();
        var path = CreateDocumentWithThreeImages("delete_last.docx");
        var outputPath = CreateTestFilePath("delete_last_out.docx");

        tool.Execute("delete", path, outputPath: outputPath, imageIndex: 2);

        var images = GetImages(tool, outputPath);
        Assert.Equal([40d, 80d], images.Select(i => i.Width));
    }

    [Fact]
    public void Replace_ShouldSwapOnlyTheAddressedImage()
    {
        var tool = new WordImageTool();
        var path = CreateDocumentWithThreeImages("replace.docx");
        var outputPath = CreateTestFilePath("replace_out.docx");
        var replacement = CreatePng("replacement.png", 300, 300);

        tool.Execute("replace", path, outputPath: outputPath, imageIndex: 0,
            imagePath: replacement, preserveSize: false);

        var images = GetImages(tool, outputPath);
        Assert.Equal(3, images.Count);
        Assert.Equal(300, images[0].OriginalSize?.WidthPixels);
        Assert.Equal(80, images[1].OriginalSize?.WidthPixels);
        Assert.Equal(120, images[2].OriginalSize?.WidthPixels);
    }
}
