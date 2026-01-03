using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptFileOperationsToolTests : TestBase
{
    private readonly PptFileOperationsTool _tool = new();

    #region General Tests

    [Fact]
    public void CreatePresentation_ShouldCreateNewPresentation()
    {
        var outputPath = CreateTestFilePath("test_create_presentation.pptx");
        _tool.Execute("create", outputPath);
        Assert.True(File.Exists(outputPath), "Presentation should be created");
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides.Count > 0, "Presentation should have at least one slide");
    }

    [Fact]
    public void ConvertPresentation_ShouldConvertToPdf()
    {
        var pptPath = CreateTestFilePath("test_convert.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "pdf");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void MergePresentations_ShouldMergePresentations()
    {
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
        _tool.Execute("merge", outputPath: outputPath, inputPaths: [ppt1Path, ppt2Path],
            keepSourceFormatting: true);
        Assert.True(File.Exists(outputPath), "Merged presentation should be created");
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides.Count >= 2, "Merged presentation should have multiple slides");
    }

    [Fact]
    public void SplitPresentation_ShouldSplitIntoMultipleFiles()
    {
        var pptPath = CreateTestFilePath("test_split.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);
        _tool.Execute("split", inputPath: pptPath, outputDirectory: outputDir);
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length >= 2, "Should create multiple files for split slides");
    }

    [Fact]
    public void ConvertPresentation_ShouldConvertToPng()
    {
        var pptPath = CreateTestFilePath("test_convert_png.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_output.png");
        var result = _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "png", slideIndex: 0);
        Assert.True(File.Exists(outputPath), "PNG file should be created");
        Assert.Contains("Slide 0", result);
        Assert.Contains("PNG", result);
    }

    [Fact]
    public void ConvertPresentation_WithSlideIndex_ShouldConvertSpecificSlide()
    {
        var pptPath = CreateTestFilePath("test_convert_slide_index.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_slide1.jpg");
        var result = _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "jpg", slideIndex: 1);
        Assert.True(File.Exists(outputPath), "JPEG file should be created");
        Assert.Contains("Slide 1", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ConvertPresentation_InvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreateTestFilePath("test_convert_invalid_slide.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_invalid.png");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "png", slideIndex: 99));
    }

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", "test.pptx"));
    }

    [Fact]
    public void MergePresentations_EmptyInputPaths_ShouldThrow()
    {
        var outputPath = CreateTestFilePath("test_merge_empty.pptx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", outputPath: outputPath, inputPaths: Array.Empty<string>()));
    }

    [Fact]
    public void SplitPresentation_InvalidSlideRange_ShouldThrow()
    {
        var pptPath = CreateTestFilePath("test_split_invalid.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "split_invalid_output");
        Assert.Throws<ArgumentException>(() => _tool.Execute("split", inputPath: pptPath, outputDirectory: outputDir,
            startSlideIndex: 10, endSlideIndex: 20));
    }

    [Fact]
    public void ConvertPresentation_UnsupportedFormat_ShouldThrow()
    {
        var pptPath = CreateTestFilePath("test_convert_unsupported.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_unsupported.xyz");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "xyz"));
    }

    #endregion

    // Note: This tool does not support session, so no Session ID Tests region
}