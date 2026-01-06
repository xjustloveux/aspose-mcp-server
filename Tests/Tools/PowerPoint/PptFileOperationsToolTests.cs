using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

[SupportedOSPlatform("windows")]
public class PptFileOperationsToolTests : TestBase
{
    private readonly PptFileOperationsTool _tool;

    public PptFileOperationsToolTests()
    {
        _tool = new PptFileOperationsTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName, int slideCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void Create_ShouldCreateNewPresentation()
    {
        var outputPath = CreateTestFilePath("test_create.pptx");
        var result = _tool.Execute("create", path: outputPath);
        Assert.StartsWith("PowerPoint presentation created successfully", result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides.Count > 0);
    }

    [Fact]
    public void Create_WithOutputPath_ShouldCreatePresentation()
    {
        var outputPath = CreateTestFilePath("test_create_output.pptx");
        var result = _tool.Execute("create", outputPath: outputPath);
        Assert.StartsWith("PowerPoint presentation created successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_ToPdf_ShouldConvert()
    {
        var pptPath = CreateTestPresentation("test_convert_pdf.pptx");
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var result = _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "pdf");
        Assert.StartsWith("Presentation from", result);
        Assert.Contains("converted to PDF format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_ToHtml_ShouldConvert()
    {
        var pptPath = CreateTestPresentation("test_convert_html.pptx");
        var outputPath = CreateTestFilePath("test_convert_output.html");
        var result = _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "html");
        Assert.StartsWith("Presentation from", result);
        Assert.Contains("converted to HTML format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_ToPng_ShouldConvertSlideToImage()
    {
        var pptPath = CreateTestPresentation("test_convert_png.pptx", 2);
        var outputPath = CreateTestFilePath("test_convert_output.png");
        var result = _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "png", slideIndex: 0);
        Assert.StartsWith("Slide 0 from", result);
        Assert.Contains("converted to PNG", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_ToJpg_WithSlideIndex_ShouldConvertSpecificSlide()
    {
        var pptPath = CreateTestPresentation("test_convert_jpg.pptx", 3);
        var outputPath = CreateTestFilePath("test_convert_slide1.jpg");
        var result = _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "jpg", slideIndex: 1);
        Assert.StartsWith("Slide 1 from", result);
        Assert.Contains("converted to JPEG", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Merge_ShouldMergePresentations()
    {
        var ppt1Path = CreateTestPresentation("test_merge1.pptx", 2);
        var ppt2Path = CreateTestPresentation("test_merge2.pptx", 2);
        var outputPath = CreateTestFilePath("test_merge_output.pptx");
        var result = _tool.Execute("merge", outputPath: outputPath, inputPaths: [ppt1Path, ppt2Path],
            keepSourceFormatting: true);
        Assert.StartsWith("Merged 2 presentations", result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        Assert.Equal(4, presentation.Slides.Count);
        foreach (var slide in presentation.Slides)
        {
            Assert.NotNull(slide.LayoutSlide);
            Assert.NotNull(slide.LayoutSlide.MasterSlide);
        }
    }

    [Fact]
    public void Merge_WithoutSourceFormatting_ShouldUseMasterFormatting()
    {
        var ppt1Path = CreateTestPresentation("test_merge_no_format1.pptx", 2);
        var ppt2Path = CreateTestPresentation("test_merge_no_format2.pptx", 2);
        var outputPath = CreateTestFilePath("test_merge_no_format_output.pptx");
        var result = _tool.Execute("merge", outputPath: outputPath, inputPaths: [ppt1Path, ppt2Path],
            keepSourceFormatting: false);
        Assert.StartsWith("Merged", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Split_ShouldSplitIntoMultipleFiles()
    {
        var pptPath = CreateTestPresentation("test_split.pptx", 3);
        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", inputPath: pptPath, outputDirectory: outputDir);
        Assert.StartsWith("Split presentation into 3 file(s)", result);
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.Equal(3, files.Length);
        foreach (var file in files)
        {
            using var presentation = new Presentation(file);
            Assert.Single(presentation.Slides);
            Assert.NotNull(presentation.Slides[0].LayoutSlide);
            Assert.NotNull(presentation.Slides[0].LayoutSlide.MasterSlide);
        }
    }

    [Fact]
    public void Split_WithSlidesPerFile_ShouldGroupSlides()
    {
        var pptPath = CreateTestPresentation("test_split_grouped.pptx", 4);
        var outputDir = Path.Combine(TestDir, "split_grouped_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", inputPath: pptPath, outputDirectory: outputDir, slidesPerFile: 2);
        Assert.StartsWith("Split presentation into 2 file(s)", result);
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.Equal(2, files.Length);
        foreach (var file in files)
        {
            using var presentation = new Presentation(file);
            Assert.Equal(2, presentation.Slides.Count);
        }
    }

    [Fact]
    public void Split_WithSlideRange_ShouldSplitSpecificRange()
    {
        var pptPath = CreateTestPresentation("test_split_range.pptx", 5);
        var outputDir = Path.Combine(TestDir, "split_range_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", inputPath: pptPath, outputDirectory: outputDir, startSlideIndex: 1,
            endSlideIndex: 3);
        Assert.StartsWith("Split presentation into 3 file(s)", result);
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.Equal(3, files.Length);
    }

    [Fact]
    public void Split_WithCustomPattern_ShouldUsePattern()
    {
        var pptPath = CreateTestPresentation("test_split_pattern.pptx", 2);
        var outputDir = Path.Combine(TestDir, "split_pattern_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", inputPath: pptPath, outputDirectory: outputDir,
            outputFileNamePattern: "page_{index}.pptx");
        Assert.StartsWith("Split presentation into", result);
        Assert.True(File.Exists(Path.Combine(outputDir, "page_0.pptx")));
    }

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive_Create(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_create_{operation}.pptx");
        var result = _tool.Execute(operation, path: outputPath);
        Assert.StartsWith("PowerPoint presentation created successfully", result);
    }

    [Theory]
    [InlineData("CONVERT")]
    [InlineData("Convert")]
    [InlineData("convert")]
    public void Operation_ShouldBeCaseInsensitive_Convert(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_convert_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_convert_{operation}_output.pdf");
        var result = _tool.Execute(operation, inputPath: pptPath, outputPath: outputPath, format: "pdf");
        Assert.StartsWith("Presentation from", result);
        Assert.Contains("converted to PDF format", result);
    }

    [Theory]
    [InlineData("MERGE")]
    [InlineData("Merge")]
    [InlineData("merge")]
    public void Operation_ShouldBeCaseInsensitive_Merge(string operation)
    {
        var ppt1Path = CreateTestPresentation($"test_case_merge1_{operation}.pptx");
        var ppt2Path = CreateTestPresentation($"test_case_merge2_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_merge_{operation}_output.pptx");
        var result = _tool.Execute(operation, outputPath: outputPath, inputPaths: [ppt1Path, ppt2Path]);
        Assert.StartsWith("Merged", result);
    }

    [Theory]
    [InlineData("SPLIT")]
    [InlineData("Split")]
    [InlineData("split")]
    public void Operation_ShouldBeCaseInsensitive_Split(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_split_{operation}.pptx", 2);
        var outputDir = Path.Combine(TestDir, $"split_case_{operation}_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute(operation, inputPath: pptPath, outputDirectory: outputDir);
        Assert.StartsWith("Split presentation into", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", path: "test.pptx"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Create_WithoutPath_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("create"));
        Assert.Contains("path or outputPath is required", ex.Message);
    }

    [Fact]
    public void Convert_WithoutInputPath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_convert_no_input.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", outputPath: outputPath, format: "pdf"));
        Assert.Contains("inputPath", ex.Message);
    }

    [Fact]
    public void Convert_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_convert_no_output.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("convert", inputPath: pptPath, format: "pdf"));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void Convert_WithoutFormat_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_convert_no_format.pptx");
        var outputPath = CreateTestFilePath("test_convert_no_format.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath));
        Assert.Contains("format is required", ex.Message);
    }

    [Fact]
    public void Convert_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_convert_unsupported.pptx");
        var outputPath = CreateTestFilePath("test_convert_unsupported.xyz");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "xyz"));
        Assert.Contains("Unsupported format", ex.Message);
    }

    [Fact]
    public void Convert_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_convert_invalid_slide.pptx");
        var outputPath = CreateTestFilePath("test_convert_invalid_slide.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "png", slideIndex: 99));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Merge_WithEmptyInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_empty.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("merge", outputPath: outputPath, inputPaths: []));
        Assert.Contains("inputPaths is required", ex.Message);
    }

    [Fact]
    public void Merge_WithNullInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_null.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("merge", outputPath: outputPath));
        Assert.Contains("inputPaths is required", ex.Message);
    }

    [Fact]
    public void Merge_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_merge_no_output.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("merge", inputPaths: [pptPath]));
        Assert.Contains("path or outputPath is required", ex.Message);
    }

    [Fact]
    public void Split_WithoutInputPath_ShouldThrowArgumentException()
    {
        var outputDir = Path.Combine(TestDir, "split_no_input");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("split", outputDirectory: outputDir));
        Assert.Contains("inputPath", ex.Message);
    }

    [Fact]
    public void Split_WithoutOutputDirectory_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_split_no_output.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("split", inputPath: pptPath));
        Assert.Contains("outputDirectory is required", ex.Message);
    }

    [Fact]
    public void Split_WithInvalidSlideRange_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_split_invalid_range.pptx", 2);
        var outputDir = Path.Combine(TestDir, "split_invalid_range");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("split", inputPath: pptPath,
            outputDirectory: outputDir, startSlideIndex: 10, endSlideIndex: 20));
        Assert.Contains("Invalid slide range", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void Convert_WithSessionId_ShouldConvertFromSession()
    {
        var pptPath = CreateTestPresentation("test_session_convert.pptx", 2);
        var sessionId = OpenSession(pptPath);
        var outputPath = CreateTestFilePath("test_session_convert_output.pdf");
        var result = _tool.Execute("convert", sessionId, outputPath: outputPath, format: "pdf");
        Assert.StartsWith("Presentation from", result);
        Assert.Contains("converted to PDF format", result);
        Assert.Contains("session", result); // Verify session was used
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_WithSessionId_ToImage_ShouldConvertSlide()
    {
        var pptPath = CreateTestPresentation("test_session_convert_img.pptx", 2);
        var sessionId = OpenSession(pptPath);
        var outputPath = CreateTestFilePath("test_session_convert_img.png");
        var result = _tool.Execute("convert", sessionId, outputPath: outputPath, format: "png", slideIndex: 1);
        Assert.StartsWith("Slide 1 from", result);
        Assert.Contains("converted to PNG", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Split_WithSessionId_ShouldSplitFromSession()
    {
        var pptPath = CreateTestPresentation("test_session_split.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var outputDir = Path.Combine(TestDir, "session_split_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", sessionId, outputDirectory: outputDir);
        Assert.StartsWith("Split presentation into 3 file(s)", result);
    }

    [Fact]
    public void Convert_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session.pdf");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("convert", "invalid_session", outputPath: outputPath, format: "pdf"));
    }

    [Fact]
    public void Split_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputDir = Path.Combine(TestDir, "invalid_session_split");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("split", "invalid_session", outputDirectory: outputDir));
    }

    [Fact]
    public void Split_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestPresentation("test_path_split.pptx");
        var pptPath2 = CreateTestPresentation("test_session_split2.pptx", 3);
        var sessionId = OpenSession(pptPath2);
        var outputDir = Path.Combine(TestDir, "prefer_session_split");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", sessionId, inputPath: pptPath1, outputDirectory: outputDir);
        Assert.Contains("3 file(s)", result);
    }

    #endregion
}