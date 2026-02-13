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

    #region Directory Creation

    [Fact]
    public void Execute_CreatesOutputDirectoryIfNotExists()
    {
        var subDir = Path.Combine(TestDir, "subdir_test", "nested");
        var outputPath = Path.Combine(subDir, "auto_created.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DirCreateTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(Directory.Exists(subDir));
        Assert.True(File.Exists(outputPath));
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

    #region Additional Barcode Types

    [Fact]
    public void Execute_QRCodeAlias_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "qrcode_alias.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "QRCodeAlias" },
            { "type", "QRCODE" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("QRCODE", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Code39Standard_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "code39standard.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CODE39TEST" },
            { "type", "Code39Standard" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("CODE39STANDARD", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Code39Extended_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "code39extended.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Code39Extended" },
            { "type", "Code39Extended" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("CODE39EXTENDED", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Ean8_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "ean8.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "12345670" },
            { "type", "EAN8" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("EAN8", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Upca_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "upca.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "012345678905" },
            { "type", "UPCA" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("UPCA", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Upce_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "upce.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "01234565" },
            { "type", "UPCE" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("UPCE", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Pdf417_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "pdf417.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "PDF417Data" },
            { "type", "PDF417" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("PDF417", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Aztec_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "aztec.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "AztecData" },
            { "type", "Aztec" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("AZTEC", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Code93_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "code93.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CODE93" },
            { "type", "Code93" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Code93Standard_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "code93standard.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CODE93STD" },
            { "type", "Code93Standard" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("CODE93STANDARD", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Code93Extended_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "code93extended.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Code93Ext" },
            { "type", "Code93Extended" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("CODE93EXTENDED", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Codabar_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "codabar.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "A12345A" },
            { "type", "Codabar" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("CODABAR", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Itf14_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "itf14.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "12345678901231" },
            { "type", "ITF14" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("ITF14", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Interleaved2of5_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "interleaved2of5.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "12345678" },
            { "type", "Interleaved2of5" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("INTERLEAVED2OF5", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Gs1Code128_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "gs1code128.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "(01)12345678901231" },
            { "type", "GS1Code128" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("GS1CODE128", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Gs1DataMatrix_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "gs1datamatrix.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "(01)12345678901231" },
            { "type", "GS1DataMatrix" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("GS1DATAMATRIX", generateResult.BarcodeType);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Additional Image Formats

    [Fact]
    public void Execute_JpegExtension_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.jpeg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "JpegExtTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("JPEG", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_GifFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.gif");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "GifTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("GIF", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_TiffFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.tiff");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "TiffTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("TIFF", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_TifFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.tif");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "TifTest" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("TIF", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_EmfFormat_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "format_test.emf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "EMFTEST" },
            { "type", "Code39" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.Equal("EMF", generateResult.ImageFormat);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Color Variations

    [Fact]
    public void Execute_WithNamedColors_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "named_colors.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "NamedColorTest" },
            { "outputPath", outputPath },
            { "foreColor", "Blue" },
            { "backColor", "Yellow" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithBothWidthAndHeight_GeneratesSuccessfully()
    {
        var outputPath = Path.Combine(TestDir, "custom_dimensions.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DimensionTest" },
            { "type", "Code128" },
            { "outputPath", outputPath },
            { "width", 4 },
            { "height", 80 }
        });

        var result = _handler.Execute(context, parameters);

        var generateResult = Assert.IsType<GenerateBarcodeResult>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(generateResult.FileSize > 0);
    }

    #endregion
}
