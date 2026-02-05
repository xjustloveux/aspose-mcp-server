using AsposeMcpServer.Results.Ocr;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Ocr;

namespace AsposeMcpServer.Tests.Tools.Ocr;

/// <summary>
///     Unit tests for OcrPreprocessingTool.
///     Tests tool construction, operation routing, and parameter validation.
/// </summary>
public class OcrPreprocessingToolTests : TestBase
{
    private readonly OcrPreprocessingTool _tool = new();

    #region Construction

    [Fact]
    public void Constructor_ShouldCreateToolWithHandlerRegistry()
    {
        var tool = new OcrPreprocessingTool();

        Assert.NotNull(tool);
    }

    #endregion

    /// <summary>
    ///     Creates a simple BMP image file for testing.
    /// </summary>
    /// <returns>The full path to the created image file.</returns>
    private string CreateTestImageFile()
    {
        var width = 10;
        var height = 10;
        var bmp = new byte[width * height * 3 + 54];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        var fileSize = bmp.Length;
        bmp[2] = (byte)(fileSize & 0xFF);
        bmp[3] = (byte)((fileSize >> 8) & 0xFF);
        bmp[4] = (byte)((fileSize >> 16) & 0xFF);
        bmp[5] = (byte)((fileSize >> 24) & 0xFF);
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = (byte)(width & 0xFF);
        bmp[19] = (byte)((width >> 8) & 0xFF);
        bmp[22] = (byte)(height & 0xFF);
        bmp[23] = (byte)((height >> 8) & 0xFF);
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
        {
            bmp[i] = 255;
            bmp[i + 1] = 0;
            bmp[i + 2] = 0;
        }

        var filePath = CreateTestFilePath($"test_image_{Guid.NewGuid()}.bmp");
        File.WriteAllBytes(filePath, bmp);
        return filePath;
    }

    #region Operation Routing

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation",
                Path.Combine(TestDir, "test.png"),
                Path.Combine(TestDir, "output.png")));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData("AUTO_SKEW")]
    [InlineData("Auto_Skew")]
    [InlineData("auto_skew")]
    public void Execute_OperationIsCaseInsensitive_AutoSkew(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, Path.Combine(TestDir, "output.png")));
    }

    [Theory]
    [InlineData("DENOISE")]
    [InlineData("Denoise")]
    [InlineData("denoise")]
    public void Execute_OperationIsCaseInsensitive_Denoise(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, Path.Combine(TestDir, "output.png")));
    }

    [Theory]
    [InlineData("CONTRAST")]
    [InlineData("Contrast")]
    [InlineData("contrast")]
    public void Execute_OperationIsCaseInsensitive_Contrast(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, Path.Combine(TestDir, "output.png")));
    }

    [Theory]
    [InlineData("SCALE")]
    [InlineData("Scale")]
    [InlineData("scale")]
    public void Execute_OperationIsCaseInsensitive_Scale(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, Path.Combine(TestDir, "output.png")));
    }

    [Theory]
    [InlineData("INVERT")]
    [InlineData("Invert")]
    [InlineData("invert")]
    public void Execute_OperationIsCaseInsensitive_Invert(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, Path.Combine(TestDir, "output.png")));
    }

    [Theory]
    [InlineData("DEWARP")]
    [InlineData("Dewarp")]
    [InlineData("dewarp")]
    public void Execute_OperationIsCaseInsensitive_Dewarp(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, Path.Combine(TestDir, "output.png")));
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("auto_skew",
                Path.Combine(TestDir, "nonexistent.png"),
                Path.Combine(TestDir, "output.png")));
    }

    [Fact]
    public void Execute_WithInvalidPath_ShouldThrowArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("auto_skew",
                "../../../etc/passwd",
                Path.Combine(TestDir, "output.png")));
    }

    #endregion

    #region Happy Path

    [SkippableFact]
    public void Execute_AutoSkew_WithValidImage_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTestImageFile();
        var outputPath = CreateTestFilePath("auto_skew_output.bmp");

        var result = _tool.Execute("auto_skew", tempImage, outputPath);

        var data = GetResultData<OcrPreprocessingResult>(result);
        Assert.Equal(tempImage, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("auto_skew", data.Operation);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_Denoise_WithValidImage_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTestImageFile();
        var outputPath = CreateTestFilePath("denoise_output.bmp");

        var result = _tool.Execute("denoise", tempImage, outputPath);

        var data = GetResultData<OcrPreprocessingResult>(result);
        Assert.Equal(tempImage, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("denoise", data.Operation);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_Contrast_WithValidImage_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTestImageFile();
        var outputPath = CreateTestFilePath("contrast_output.bmp");

        var result = _tool.Execute("contrast", tempImage, outputPath);

        var data = GetResultData<OcrPreprocessingResult>(result);
        Assert.Equal(tempImage, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("contrast", data.Operation);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_Scale_WithCustomFactor_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTestImageFile();
        var outputPath = CreateTestFilePath("scale_output.bmp");

        var result = _tool.Execute("scale", tempImage, outputPath, 3.0);

        var data = GetResultData<OcrPreprocessingResult>(result);
        Assert.Equal(tempImage, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("scale", data.Operation);
        Assert.Contains("3.0", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_Invert_WithValidImage_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTestImageFile();
        var outputPath = CreateTestFilePath("invert_output.bmp");

        var result = _tool.Execute("invert", tempImage, outputPath);

        var data = GetResultData<OcrPreprocessingResult>(result);
        Assert.Equal(tempImage, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("invert", data.Operation);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_Dewarp_WithValidImage_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTestImageFile();
        var outputPath = CreateTestFilePath("dewarp_output.bmp");

        var result = _tool.Execute("dewarp", tempImage, outputPath);

        var data = GetResultData<OcrPreprocessingResult>(result);
        Assert.Equal(tempImage, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("dewarp", data.Operation);
        Assert.True(File.Exists(outputPath));
    }

    #endregion
}
