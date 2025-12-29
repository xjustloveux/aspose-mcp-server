using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptFileOperationsToolTests : TestBase
{
    private readonly PptFileOperationsTool _tool = new();

    [Fact]
    public async Task CreatePresentation_ShouldCreateNewPresentation()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_presentation.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "create",
            ["path"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Presentation should be created");
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides.Count > 0, "Presentation should have at least one slide");
    }

    [Fact]
    public async Task ConvertPresentation_ShouldConvertToPdf()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath,
            ["format"] = "pdf"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task MergePresentations_ShouldMergePresentations()
    {
        // Arrange
        var ppt1Path = CreateTestFilePath("test_merge1.pptx");
        using (var ppt1 = new Presentation())
        {
            ppt1.Slides.AddEmptySlide(ppt1.LayoutSlides[0]);
            ppt1.Save(ppt1Path, SaveFormat.Pptx);
        }

        var ppt2Path = CreateTestFilePath("test_merge2.pptx");
        using (var ppt2 = new Presentation())
        {
            ppt2.Slides.AddEmptySlide(ppt2.LayoutSlides[0]);
            ppt2.Save(ppt2Path, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_merge_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["inputPath"] = ppt1Path,
            ["outputPath"] = outputPath,
            ["inputPaths"] = new JsonArray { ppt2Path },
            ["keepSourceFormatting"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Merged presentation should be created");
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides.Count >= 2, "Merged presentation should have multiple slides");
    }

    [Fact]
    public async Task SplitPresentation_ShouldSplitIntoMultipleFiles()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_split.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["inputPath"] = pptPath,
            ["outputDirectory"] = outputDir
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length >= 2, "Should create multiple files for split slides");
    }

    [Fact]
    public async Task ConvertPresentation_ShouldConvertToPng()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert_png.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_output.png");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath,
            ["format"] = "png",
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PNG file should be created");
        Assert.Contains("Slide 0", result);
        Assert.Contains("PNG", result);
    }

    [Fact]
    public async Task ConvertPresentation_WithSlideIndex_ShouldConvertSpecificSlide()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert_slide_index.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_slide1.jpg");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath,
            ["format"] = "jpg",
            ["slideIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "JPEG file should be created");
        Assert.Contains("Slide 1", result);
    }

    [Fact]
    public async Task ConvertPresentation_InvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert_invalid_slide.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_invalid.png");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath,
            ["format"] = "png",
            ["slideIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = "test.pptx"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task MergePresentations_EmptyInputPaths_ShouldThrow()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_merge_empty.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["outputPath"] = outputPath,
            ["inputPaths"] = new JsonArray()
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SplitPresentation_InvalidSlideRange_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_split_invalid.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "split_invalid_output");
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["inputPath"] = pptPath,
            ["outputDirectory"] = outputDir,
            ["startSlideIndex"] = 10,
            ["endSlideIndex"] = 20
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ConvertPresentation_UnsupportedFormat_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert_unsupported.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_unsupported.xyz");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath,
            ["format"] = "xyz"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}