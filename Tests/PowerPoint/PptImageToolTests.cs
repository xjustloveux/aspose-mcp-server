using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

[SupportedOSPlatform("windows")]
public class PptImageToolTests : TestBase
{
    private readonly PptImageTool _tool = new();

    private string CreateTestImage(string fileName, int width = 10, int height = 10, Color? color = null)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(width, height);
        var fillColor = color ?? Color.Red;
        for (var x = 0; x < width; x++)
        for (var y = 0; y < height; y++)
            bitmap.SetPixel(x, y, fillColor);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    private string CreateTestPresentation(string fileName, int slideCount = 1, bool addImages = false)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();

        for (var i = 0; i < slideCount; i++)
        {
            var slide = i == 0
                ? presentation.Slides[0]
                : presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            if (addImages)
            {
                var imagePath = CreateTestImage($"pres_image_{i}.png");
                using var imageStream = File.OpenRead(imagePath);
                var pictureImage = presentation.Images.AddImage(imageStream);
                slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 100, 100, 200, 150, pictureImage);
            }
        }

        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region Unknown Operation Tests

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region Add Operation Tests

    [Fact]
    public async Task Add_ShouldAddImageToSlide()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_image.pptx");
        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imagePath"] = imagePath,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 150
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image added", result);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var images = slide.Shapes.OfType<IPictureFrame>().ToList();
        Assert.True(images.Count > 0, "Slide should contain at least one image");
    }

    [Fact]
    public async Task Add_WithOnlyWidth_ShouldMaintainAspectRatio()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_aspect.pptx");
        var imagePath = CreateTestImage("test_image_aspect.png", 100, 50);
        var outputPath = CreateTestFilePath("test_add_aspect_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imagePath"] = imagePath,
            ["width"] = 200
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var image = presentation.Slides[0].Shapes.OfType<IPictureFrame>().First();
        Assert.Equal(200, image.Width);
        Assert.Equal(100, image.Height); // 200 * (50/100) = 100
    }

    #endregion

    #region Edit Operation Tests

    [Fact]
    public async Task Edit_ShouldModifyImageSize()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_image.pptx", addImages: true);
        var outputPath = CreateTestFilePath("test_edit_image_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0,
            ["width"] = 300,
            ["height"] = 200
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("updated", result);
        using var presentation = new Presentation(outputPath);
        var images = presentation.Slides[0].Shapes.OfType<IPictureFrame>().ToList();
        Assert.True(images.Count > 0);
        Assert.Equal(300, images[0].Width);
        Assert.Equal(200, images[0].Height);
    }

    [Fact]
    public async Task Edit_WithNewImage_ShouldReplaceImage()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_replace.pptx", addImages: true);
        var newImagePath = CreateTestImage("new_image.png");
        var outputPath = CreateTestFilePath("test_edit_replace_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0,
            ["imagePath"] = newImagePath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("image replaced", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Edit_WithJpegQuality_ShouldCompressImage()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_compress.pptx", addImages: true);
        var newImagePath = CreateTestImage("compress_image.png");
        var outputPath = CreateTestFilePath("test_edit_compress_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0,
            ["imagePath"] = newImagePath,
            ["jpegQuality"] = 50
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("quality=50", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Edit_WithMaxWidth_ShouldResizeImage()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_resize.pptx", addImages: true);
        var largeImagePath = CreateTestImage("large_image.png", 200, 100, Color.Blue);
        var outputPath = CreateTestFilePath("test_edit_resize_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0,
            ["imagePath"] = largeImagePath,
            ["maxWidth"] = 100
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("resized", result);
        Assert.Contains("100x50", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Edit_WithMaxHeight_ShouldResizeImage()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_resize_height.pptx", addImages: true);
        var largeImagePath = CreateTestImage("large_height.png", 100, 200, Color.Green);
        var outputPath = CreateTestFilePath("test_edit_resize_height_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0,
            ["imagePath"] = largeImagePath,
            ["maxHeight"] = 100
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("resized", result);
        Assert.Contains("50x100", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Edit_WithMaxWidthAndQuality_ShouldResizeAndCompress()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_both.pptx", addImages: true);
        var largeImagePath = CreateTestImage("large_both.png", 200, 200, Color.Yellow);
        var outputPath = CreateTestFilePath("test_edit_both_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0,
            ["imagePath"] = largeImagePath,
            ["maxWidth"] = 100,
            ["jpegQuality"] = 75
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("resized", result);
        Assert.Contains("quality=75", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Edit_ImageSmallerThanMax_ShouldNotResize()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_no_resize.pptx", addImages: true);
        var smallImagePath = CreateTestImage("small_image.png"); // 10x10
        var outputPath = CreateTestFilePath("test_edit_no_resize_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0,
            ["imagePath"] = smallImagePath,
            ["maxWidth"] = 100,
            ["maxHeight"] = 100
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.DoesNotContain("resized", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Edit_WithInvalidImageIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_invalid.pptx", addImages: true);
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 99,
            ["width"] = 300
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("imageIndex", ex.Message);
    }

    #endregion

    #region Delete Operation Tests

    [Fact]
    public async Task Delete_ShouldRemoveImageFromSlide()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete.pptx", addImages: true);
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("deleted", result);
        using var presentation = new Presentation(outputPath);
        var images = presentation.Slides[0].Shapes.OfType<IPictureFrame>().ToList();
        Assert.Empty(images);
    }

    [Fact]
    public async Task Delete_WithInvalidImageIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_invalid.pptx", addImages: true);
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["imageIndex"] = 99
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("imageIndex", ex.Message);
    }

    #endregion

    #region Get Operation Tests

    [Fact]
    public async Task Get_ShouldReturnImageInfo()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get.pptx", addImages: true);
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.Equal(1, json.RootElement.GetProperty("imageCount").GetInt32());
        Assert.True(json.RootElement.GetProperty("images").GetArrayLength() > 0);
    }

    [Fact]
    public async Task Get_EmptySlide_ShouldReturnEmptyImages()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_empty.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("imageCount").GetInt32());
    }

    #endregion

    #region Export Slides Operation Tests

    [Fact]
    public async Task ExportSlides_ShouldExportAllSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_export.pptx", 3, true);
        var outputDir = Path.Combine(TestDir, "exported_slides");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "export_slides",
            ["path"] = pptPath,
            ["outputDir"] = outputDir,
            ["format"] = "png"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir, "*.png");
        Assert.Equal(3, files.Length);
        Assert.Contains("Exported 3 slides", result);
    }

    [Fact]
    public async Task ExportSlides_WithSlideIndexes_ShouldExportSpecificSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_export_specific.pptx", 5, true);
        var outputDir = Path.Combine(TestDir, "exported_specific");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "export_slides",
            ["path"] = pptPath,
            ["outputDir"] = outputDir,
            ["slideIndexes"] = "0,2,4"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir, "*.png");
        Assert.Equal(3, files.Length);
        Assert.Contains("Exported 3 slides", result);
        Assert.True(File.Exists(Path.Combine(outputDir, "slide_1.png")));
        Assert.True(File.Exists(Path.Combine(outputDir, "slide_3.png")));
        Assert.True(File.Exists(Path.Combine(outputDir, "slide_5.png")));
    }

    [Fact]
    public async Task ExportSlides_WithInvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_export_invalid.pptx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "export_slides",
            ["path"] = pptPath,
            ["slideIndexes"] = "0,10"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("slideIndex 10", ex.Message);
    }

    [Fact]
    public async Task ExportSlides_WithInvalidSlideIndexFormat_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_export_invalid_format.pptx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "export_slides",
            ["path"] = pptPath,
            ["slideIndexes"] = "0,abc,2"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid slide index", ex.Message);
    }

    [Fact]
    public async Task ExportSlides_WithJpegFormat_ShouldExportAsJpeg()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_export_jpeg.pptx", 2);
        var outputDir = Path.Combine(TestDir, "exported_jpeg");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "export_slides",
            ["path"] = pptPath,
            ["outputDir"] = outputDir,
            ["format"] = "jpeg"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir, "*.jpg");
        Assert.Equal(2, files.Length);
    }

    [Fact]
    public async Task ExportSlides_WithScale_ShouldApplyScaling()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_export_scale.pptx");
        var outputDir = Path.Combine(TestDir, "exported_scale");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "export_slides",
            ["path"] = pptPath,
            ["outputDir"] = outputDir,
            ["scale"] = 0.5
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Exported 1 slides", result);
        var files = Directory.GetFiles(outputDir, "*.png");
        Assert.Single(files);
    }

    #endregion

    #region Extract Operation Tests

    [Fact]
    public async Task Extract_ShouldExtractImages()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_extract.pptx", 2, true);
        var outputDir = Path.Combine(TestDir, "extracted_images");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = pptPath,
            ["outputDir"] = outputDir
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length > 0, "Should extract at least one image");
        Assert.Contains("Extracted", result);
    }

    [Fact]
    public async Task Extract_WithSkipDuplicates_ShouldSkipDuplicateImages()
    {
        // Arrange - Create presentation with duplicate images
        var filePath = CreateTestFilePath("test_extract_duplicates.pptx");
        var imagePath = CreateTestImage("duplicate_image.png");

        using (var presentation = new Presentation())
        {
            await using var imageStream = File.OpenRead(imagePath);
            var pictureImage = presentation.Images.AddImage(imageStream);

            // Add same image to multiple slides
            for (var i = 0; i < 3; i++)
            {
                var slide = i == 0
                    ? presentation.Slides[0]
                    : presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
                slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 100, 100, 200, 150, pictureImage);
            }

            presentation.Save(filePath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "extracted_skip_duplicates");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = filePath,
            ["outputDir"] = outputDir,
            ["skipDuplicates"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert - Should only export 1 image (others are duplicates)
        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
        Assert.Contains("skipped", result);
        Assert.Contains("duplicates", result);
    }

    [Fact]
    public async Task Extract_WithoutSkipDuplicates_ShouldExtractAllImages()
    {
        // Arrange - Create presentation with duplicate images
        var filePath = CreateTestFilePath("test_extract_all.pptx");
        var imagePath = CreateTestImage("all_images.png");

        using (var presentation = new Presentation())
        {
            await using var imageStream = File.OpenRead(imagePath);
            var pictureImage = presentation.Images.AddImage(imageStream);

            // Add same image to multiple slides
            for (var i = 0; i < 3; i++)
            {
                var slide = i == 0
                    ? presentation.Slides[0]
                    : presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
                slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 100, 100, 200, 150, pictureImage);
            }

            presentation.Save(filePath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "extracted_all");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = filePath,
            ["outputDir"] = outputDir,
            ["skipDuplicates"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert - Should export all 3 images
        var files = Directory.GetFiles(outputDir);
        Assert.Equal(3, files.Length);
        Assert.DoesNotContain("skipped", result);
    }

    #endregion
}