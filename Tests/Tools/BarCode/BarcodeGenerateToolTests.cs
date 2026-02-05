using AsposeMcpServer.Results;
using AsposeMcpServer.Results.BarCode.Generate;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.BarCode;

namespace AsposeMcpServer.Tests.Tools.BarCode;

/// <summary>
///     Integration tests for <see cref="BarcodeGenerateTool" />.
///     Focuses on operation routing, image generation, and end-to-end barcode generation operations.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class BarcodeGenerateToolTests : BarcodeTestBase
{
    private readonly BarcodeGenerateTool _tool = new();

    #region Generate Operations

    [Fact]
    public void Generate_QrCode_DefaultType_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("qr_default.png");

        var result = _tool.Execute("generate", "Hello World", outputPath);

        Assert.NotNull(result);
        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("QR", data.BarcodeType);
        Assert.Equal("Hello World", data.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_Code128_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("code128.png");

        var result = _tool.Execute("generate", "12345", outputPath, "Code128");

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("CODE128", data.BarcodeType);
        Assert.Equal("12345", data.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_Ean13_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("ean13.png");

        var result = _tool.Execute("generate", "5901234123457", outputPath, "EAN13");

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("EAN13", data.BarcodeType);
        Assert.Equal("5901234123457", data.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_DataMatrix_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("datamatrix.png");

        var result = _tool.Execute("generate", "DataMatrixTest", outputPath, "DataMatrix");

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("DATAMATRIX", data.BarcodeType);
        Assert.Equal("DataMatrixTest", data.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_WithCustomType_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("code39.png");

        var result = _tool.Execute("generate", "ABCDE", outputPath, "Code39");

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("CODE39", data.BarcodeType);
        Assert.Equal("ABCDE", data.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Image Format Tests

    [Fact]
    public void Generate_PngFormat_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("format_test.png");

        var result = _tool.Execute("generate", "PNG Test", outputPath);

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("PNG", data.ImageFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Generate_JpegFormat_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("format_test.jpg");

        var result = _tool.Execute("generate", "JPEG Test", outputPath);

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("JPG", data.ImageFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Generate_BmpFormat_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("format_test.bmp");

        var result = _tool.Execute("generate", "BMP Test", outputPath);

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("BMP", data.ImageFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Generate_SvgFormat_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("format_test.svg");

        var result = _tool.Execute("generate", "SVGTEST", outputPath, "Code39");

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("SVG", data.ImageFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region Custom Parameters

    [Fact]
    public void Generate_WithWidth_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("width_test.png");

        var result = _tool.Execute("generate", "Width Test", outputPath, width: 5);

        Assert.IsType<FinalizedResult<GenerateBarcodeResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Generate_WithHeight_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("height_test.png");

        var result = _tool.Execute("generate", "Height Test", outputPath, height: 100);

        Assert.IsType<FinalizedResult<GenerateBarcodeResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Generate_WithForeColor_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("forecolor_test.png");

        var result = _tool.Execute("generate", "ForeColor", outputPath, foreColor: "#FF0000");

        Assert.IsType<FinalizedResult<GenerateBarcodeResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Generate_WithBackColor_ShouldGenerate()
    {
        var outputPath = CreateTestFilePath("backcolor_test.png");

        var result = _tool.Execute("generate", "BackColor", outputPath, backColor: "#FFFF00");

        Assert.IsType<FinalizedResult<GenerateBarcodeResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region Result Validation

    [Fact]
    public void Generate_ResultContainsCorrectOutputPath()
    {
        var outputPath = CreateTestFilePath("result_path.png");

        var result = _tool.Execute("generate", "PathTest", outputPath);

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal(outputPath, data.OutputPath);
    }

    [Fact]
    public void Generate_ResultContainsCorrectBarcodeType()
    {
        var outputPath = CreateTestFilePath("result_type.png");

        var result = _tool.Execute("generate", "TypeTest", outputPath, "Code128");

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal("CODE128", data.BarcodeType);
    }

    [Fact]
    public void Generate_ResultContainsEncodedText()
    {
        var outputPath = CreateTestFilePath("result_text.png");
        var expectedText = "EncodedTextTest";

        var result = _tool.Execute("generate", expectedText, outputPath);

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Equal(expectedText, data.EncodedText);
    }

    [Fact]
    public void Generate_ResultHasPositiveFileSize()
    {
        var outputPath = CreateTestFilePath("result_filesize.png");

        var result = _tool.Execute("generate", "FileSizeTest", outputPath);

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.NotNull(data.FileSize);
        Assert.True(data.FileSize > 0);
    }

    [Fact]
    public void Generate_ResultMessageContainsGenerated()
    {
        var outputPath = CreateTestFilePath("result_message.png");

        var result = _tool.Execute("generate", "MessageTest", outputPath);

        var data = GetResultData<GenerateBarcodeResult>(result);
        Assert.Contains("generated", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Generate_OutputFileExistsAndNonEmpty()
    {
        var outputPath = CreateTestFilePath("result_file.png");

        _tool.Execute("generate", "FileTest", outputPath);

        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GENERATE")]
    [InlineData("Generate")]
    [InlineData("generate")]
    public void Execute_CaseInsensitiveOperation_ShouldWork(string operation)
    {
        var outputPath = CreateTestFilePath($"case_{operation}.png");

        var result = _tool.Execute(operation, "CaseTest", outputPath);

        Assert.NotNull(result);
        Assert.IsType<FinalizedResult<GenerateBarcodeResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_UnknownOperation_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("unknown_op.png");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", "Test", outputPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Generate_MissingText_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("missing_text.png");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("generate", outputPath: outputPath));
    }

    [Fact]
    public void Generate_MissingOutputPath_ShouldThrowArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("generate", "NoOutput"));
    }

    [Fact]
    public void Generate_UnsupportedImageFormat_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("unsupported.xyz");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("generate", "FormatTest", outputPath));
        Assert.Contains("Unsupported image format", ex.Message);
    }

    #endregion
}
