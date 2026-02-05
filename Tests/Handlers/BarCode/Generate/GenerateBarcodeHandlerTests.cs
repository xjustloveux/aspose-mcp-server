using AsposeMcpServer.Handlers.BarCode.Generate;
using AsposeMcpServer.Results.BarCode.Generate;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.BarCode.Generate;

/// <summary>
///     Tests for <see cref="GenerateBarcodeHandler" />.
///     Verifies barcode generation with various types, image formats, and custom parameters.
/// </summary>
public class GenerateBarcodeHandlerTests : HandlerTestBase<object>
{
    private readonly GenerateBarcodeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Generate()
    {
        Assert.Equal("generate", _handler.Operation);
    }

    #endregion

    #region Basic Generation

    [Fact]
    public void Execute_QrCode_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "qr_default.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "TestQRData" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("QR", generateResult.BarcodeType);
        Assert.Equal("TestQRData", generateResult.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Code128_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "code128.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Code128Data" },
            { "type", "Code128" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("CODE128", generateResult.BarcodeType);
        Assert.Equal("Code128Data", generateResult.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Ean13_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "ean13.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "5901234123457" },
            { "type", "EAN13" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("EAN13", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_DataMatrix_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "datamatrix.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DataMatrixTest" },
            { "type", "DataMatrix" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("DATAMATRIX", generateResult.BarcodeType);
        Assert.Equal("DataMatrixTest", generateResult.EncodedText);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Image Formats

    [Fact]
    public void Execute_PngFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "PngTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("PNG", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_JpegFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.jpg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "JpegTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("JPG", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_BmpFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.bmp");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "BmpTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("BMP", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_SvgFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.svg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "SvgTest" },
            { "type", "Code39" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("SVG", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Custom Parameters

    [Fact]
    public void Execute_WithCustomWidth_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "custom_width.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "WidthTest" },
            { "outputPath", outputPath },
            { "width", 5 }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(File.Exists(outputPath));
        Assert.NotNull(generateResult.FileSize);
        Assert.True(generateResult.FileSize > 0);
    }

    [Fact]
    public void Execute_WithCustomHeight_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "custom_height.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "HeightTest" },
            { "type", "Code128" },
            { "outputPath", outputPath },
            { "height", 100 }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(File.Exists(outputPath));
        Assert.NotNull(generateResult.FileSize);
        Assert.True(generateResult.FileSize > 0);
    }

    [Fact]
    public void Execute_WithCustomForeColor_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "custom_forecolor.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "ForeColorTest" },
            { "outputPath", outputPath },
            { "foreColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(File.Exists(outputPath));
        Assert.NotNull(generateResult.FileSize);
        Assert.True(generateResult.FileSize > 0);
    }

    [Fact]
    public void Execute_WithCustomBackColor_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "custom_backcolor.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "BackColorTest" },
            { "outputPath", outputPath },
            { "backColor", "#FFFF00" }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(File.Exists(outputPath));
        Assert.NotNull(generateResult.FileSize);
        Assert.True(generateResult.FileSize > 0);
    }

    #endregion

    #region Result Validation

    [Fact]
    public void Execute_ResultHasCorrectBarcodeType()
    {
        var outputPath = Path.Combine(TestDir, "result_type.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "TypeCheck" },
            { "type", "Code128" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("CODE128", generateResult.BarcodeType);
    }

    [Fact]
    public void Execute_ResultHasCorrectEncodedText()
    {
        var outputPath = Path.Combine(TestDir, "result_text.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "EncodedTextCheck" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("EncodedTextCheck", generateResult.EncodedText);
    }

    [Fact]
    public void Execute_ResultHasCorrectImageFormat()
    {
        var outputPath = Path.Combine(TestDir, "result_format.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "FormatCheck" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("PNG", generateResult.ImageFormat);
    }

    [Fact]
    public void Execute_ResultHasPositiveFileSize()
    {
        var outputPath = Path.Combine(TestDir, "result_filesize.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "FileSizeCheck" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.NotNull(generateResult.FileSize);
        Assert.True(generateResult.FileSize > 0);
    }

    [Fact]
    public void Execute_ResultMessageContainsGenerated()
    {
        var outputPath = Path.Combine(TestDir, "result_message.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "MessageCheck" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Contains("generated", generateResult.Message);
    }

    [Fact]
    public void Execute_ResultHasCorrectOutputPath()
    {
        var outputPath = Path.Combine(TestDir, "result_outputpath.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "OutputPathCheck" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal(outputPath, generateResult.OutputPath);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingText_ThrowsArgumentException()
    {
        var outputPath = Path.Combine(TestDir, "error_no_text.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingOutputPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "NoOutputPath" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedBarcodeType_ThrowsArgumentException()
    {
        var outputPath = Path.Combine(TestDir, "error_bad_type.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "BadType" },
            { "type", "INVALID_TYPE" },
            { "outputPath", outputPath }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported barcode type", ex.Message);
    }

    [Fact]
    public void Execute_WithUnsupportedImageFormat_ThrowsArgumentException()
    {
        var outputPath = Path.Combine(TestDir, "error_bad_format.xyz");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "BadFormat" },
            { "outputPath", outputPath }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported image format", ex.Message);
    }

    #endregion
}
