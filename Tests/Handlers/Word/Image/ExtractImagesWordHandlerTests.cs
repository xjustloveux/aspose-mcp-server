using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Image;

public class ExtractImagesWordHandlerTests : WordHandlerTestBase
{
    private readonly ExtractImagesWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Extract()
    {
        Assert.Equal("extract", _handler.Operation);
    }

    #endregion

    #region Basic Extract Operations

    [Fact]
    public void Execute_WithNoImages_ReturnsNoImagesMessage()
    {
        var outputDir = Path.Combine(TestDir, "extract_output");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("No images found", result.Message);
    }

    [Fact]
    public void Execute_ExtractsImages()
    {
        var tempImageFile = CreateTempImageFile();
        var outputDir = Path.Combine(TestDir, "extract_output");
        var doc = CreateDocumentWithImage(tempImageFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Successfully extracted", result.Message);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir);
        Assert.NotEmpty(files);
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            Assert.True(fileInfo.Length > 0, $"Extracted image {file} should have content");
        }
    }

    [Fact]
    public void Execute_ReturnsImageCount()
    {
        var tempImageFile = CreateTempImageFile();
        var outputDir = Path.Combine(TestDir, "extract_count");
        var doc = CreateDocumentWithMultipleImages(tempImageFile, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3 images", result.Message);
    }

    /// <summary>
    ///     R5-R01: extraction published each image as it was produced, so a request that failed on a
    ///     later image had already replaced the destinations of the earlier ones. The publish for the
    ///     second image is made to fail by putting a directory where its file has to go.
    /// </summary>
    [Fact]
    public void Execute_WhenALaterImageCannotBePublished_ShouldWriteNoneOfThem()
    {
        var tempImageFile = CreateTempImageFile();
        var outputDir = Path.Combine(TestDir, "atomic_images");
        Directory.CreateDirectory(outputDir);

        var doc = CreateDocumentWithMultipleImages(tempImageFile, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "prefix", "img" }
        });

        // The handler names each file from the image's own type, not from the fixture file's
        // extension, so the names are learned from a successful run rather than guessed.
        var probeDir = Path.Combine(TestDir, "atomic_images_probe");
        _handler.Execute(CreateContext(CreateDocumentWithMultipleImages(tempImageFile, 3)),
            CreateParameters(new Dictionary<string, object?>
            {
                { "outputDir", probeDir },
                { "prefix", "img" }
            }));
        var produced = Directory.GetFiles(probeDir).Select(Path.GetFileName).Order(StringComparer.Ordinal).ToList();
        Assert.Equal(3, produced.Count);

        var first = Path.Combine(outputDir, produced[0]!);
        Directory.CreateDirectory(Path.Combine(outputDir, produced[1]!));

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));

        Assert.False(System.IO.File.Exists(first),
            "the first image was published before the second was known to fail");
        Assert.DoesNotContain(Directory.GetFiles(outputDir),
            f => Path.GetFileName(f).Contains(".partial-", StringComparison.Ordinal));
    }

    #endregion

    #region Optional Parameters

    [Fact]
    public void Execute_WithPrefix_UsesPrefix()
    {
        var tempImageFile = CreateTempImageFile();
        var outputDir = Path.Combine(TestDir, "extract_prefix");
        var doc = CreateDocumentWithImage(tempImageFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "prefix", "myimage" }
        });

        _handler.Execute(context, parameters);

        var files = Directory.GetFiles(outputDir);
        Assert.NotEmpty(files);
        Assert.Contains(files, f => Path.GetFileName(f).StartsWith("myimage_"));
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            Assert.True(fileInfo.Length > 0, $"Extracted image {file} should have content");
        }
    }

    [Fact]
    public void Execute_WithExtractImageIndex_ExtractsSingleImage()
    {
        var tempImageFile = CreateTempImageFile();
        var outputDir = Path.Combine(TestDir, "extract_single");
        var doc = CreateDocumentWithMultipleImages(tempImageFile, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "extractImageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("image #1", result.Message);
        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
        var fileInfo = new FileInfo(files[0]);
        Assert.True(fileInfo.Length > 0, "Extracted image should have content");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputDir_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("outputDir", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidExtractImageIndex_ThrowsArgumentException()
    {
        var tempImageFile = CreateTempImageFile();
        var outputDir = Path.Combine(TestDir, "extract_invalid");
        var doc = CreateDocumentWithImage(tempImageFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "extractImageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithImage(string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        return doc;
    }

    private static Document CreateDocumentWithMultipleImages(string imagePath, int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++)
        {
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
        }

        return doc;
    }

    #endregion
}
