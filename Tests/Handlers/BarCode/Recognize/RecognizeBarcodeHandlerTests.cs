using System.Drawing;
using System.Drawing.Imaging;
using Aspose.BarCode.Generation;
using AsposeMcpServer.Handlers.BarCode.Recognize;
using AsposeMcpServer.Results.BarCode.Recognize;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.BarCode.Recognize;

/// <summary>
///     Tests for <see cref="RecognizeBarcodeHandler" />.
///     Verifies barcode recognition from images with various barcode types and decode filters.
/// </summary>
public class RecognizeBarcodeHandlerTests : HandlerTestBase<object>
{
    private readonly RecognizeBarcodeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Recognize()
    {
        Assert.Equal("recognize", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a barcode image file for testing recognition.
    /// </summary>
    /// <param name="fileName">The output image file name.</param>
    /// <param name="text">The text to encode in the barcode.</param>
    /// <param name="encodeType">The barcode type (default: QR).</param>
    /// <returns>The full path to the created barcode image.</returns>
    private string CreateTestBarcodeImage(string fileName, string text = "TestData",
        BaseEncodeType? encodeType = null)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var generator = new BarcodeGenerator(encodeType ?? EncodeTypes.QR, text);
        generator.Save(filePath, BarCodeImageFormat.Png);
        return filePath;
    }

    #endregion

    #region Additional Helper Methods

    private string CreateEmptyImage(string fileName)
    {
        var filePath = Path.Combine(TestDir, fileName);
        using var bitmap = new Bitmap(100, 100);
        using var g = Graphics.FromImage(bitmap);
        g.Clear(Color.White);
        bitmap.Save(filePath, ImageFormat.Png);
        return filePath;
    }

    #endregion

    #region Basic Recognition

    [Fact]
    public void Execute_RecognizeQrCode_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_qr.png", "QRTestData", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
        Assert.NotEmpty(recognizeResult.Barcodes);
    }

    [Fact]
    public void Execute_RecognizeCode128_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_code128.png", "Code128TestData", EncodeTypes.Code128);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
        Assert.NotEmpty(recognizeResult.Barcodes);
    }

    [Fact]
    public void Execute_WithSpecificTypeFilter_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_qr_filter.png", "FilterTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "QR" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
        Assert.NotEmpty(recognizeResult.Barcodes);
    }

    [Fact]
    public void Execute_WithAutoType_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_auto.png", "AutoTypeTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "auto" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
        Assert.NotEmpty(recognizeResult.Barcodes);
    }

    #endregion

    #region Result Validation

    [Fact]
    public void Execute_ResultHasCorrectSourcePath()
    {
        var imagePath = CreateTestBarcodeImage("result_source.png", "SourcePathTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.Equal(imagePath, recognizeResult.SourcePath);
    }

    [Fact]
    public void Execute_ResultHasPositiveCount()
    {
        var imagePath = CreateTestBarcodeImage("result_count.png", "CountTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_ResultBarcodesListIsNotEmpty()
    {
        var imagePath = CreateTestBarcodeImage("result_list.png", "ListTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.NotEmpty(recognizeResult.Barcodes);
    }

    [Fact]
    public void Execute_EachBarcodeInfoHasNonEmptyCodeText()
    {
        var imagePath = CreateTestBarcodeImage("result_codetext.png", "CodeTextTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.NotEmpty(recognizeResult.Barcodes);
        foreach (var barcode in recognizeResult.Barcodes)
            Assert.False(string.IsNullOrEmpty(barcode.CodeText));
    }

    [Fact]
    public void Execute_EachBarcodeInfoHasNonEmptyCodeType()
    {
        var imagePath = CreateTestBarcodeImage("result_codetype.png", "CodeTypeTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.NotEmpty(recognizeResult.Barcodes);
        foreach (var barcode in recognizeResult.Barcodes)
            Assert.False(string.IsNullOrEmpty(barcode.CodeType));
    }

    [Fact]
    public void Execute_ResultMessageContainsFound()
    {
        var imagePath = CreateTestBarcodeImage("result_message.png", "MessageTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.Contains("Found", recognizeResult.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>());

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var nonExistentPath = Path.Combine(TestDir, "nonexistent_barcode.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", nonExistentPath }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedDecodeType_ThrowsArgumentException()
    {
        var imagePath = CreateTestBarcodeImage("error_decode_type.png", "DecodeTypeTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "INVALID_DECODE_TYPE" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported decode type", ex.Message);
    }

    #endregion

    #region Decode Type Aliases

    [Fact]
    public void Execute_WithAllType_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_all.png", "AllTypeTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "all" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
        Assert.Equal("All", recognizeResult.DecodeType);
    }

    [Fact]
    public void Execute_WithAllSupportedTypes_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_allsupported.png", "AllSupportedTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "AllSupportedTypes" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
        Assert.Equal("All", recognizeResult.DecodeType);
    }

    [Fact]
    public void Execute_WithQRCodeAlias_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_qrcode_alias.png", "QRCodeAliasTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "QRCODE" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_WithCode39Standard_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_code39std.png", "CODE39TEST", EncodeTypes.Code39Standard);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Code39Standard" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_WithCode39Alias_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_code39.png", "CODE39DATA", EncodeTypes.Code39Standard);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Code39" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    #endregion

    #region Additional Barcode Type Recognition

    [Fact]
    public void Execute_RecognizeEan13_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_ean13.png", "5901234123457", EncodeTypes.EAN13);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "EAN13" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeEan8_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_ean8.png", "12345670", EncodeTypes.EAN8);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "EAN8" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeUpca_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_upca.png", "012345678905", EncodeTypes.UPCA);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "UPCA" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeUpce_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_upce.png", "01234565", EncodeTypes.UPCE);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "UPCE" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeDataMatrix_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_datamatrix.png", "DataMatrixTest", EncodeTypes.DataMatrix);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "DataMatrix" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizePdf417_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_pdf417.png", "PDF417Data", EncodeTypes.Pdf417);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "PDF417" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeCode39Standard_WithExplicitType_ReturnsResults()
    {
        var imagePath =
            CreateTestBarcodeImage("recognize_code39std_explicit.png", "CODE39DATA", EncodeTypes.Code39Standard);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Code39Standard" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeCode93_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_code93.png", "CODE93", EncodeTypes.Code93Standard);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Code93" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeCode93Standard_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_code93std.png", "CODE93STD", EncodeTypes.Code93Standard);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Code93Standard" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeCodabar_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_codabar.png", "A12345A", EncodeTypes.Codabar);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Codabar" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeItf14_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_itf14.png", "12345678901231", EncodeTypes.ITF14);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "ITF14" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeInterleaved2of5_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_interleaved.png", "12345678", EncodeTypes.Interleaved2of5);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Interleaved2of5" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeGs1Code128_ReturnsResults()
    {
        var imagePath =
            CreateTestBarcodeImage("recognize_gs1code128.png", "(01)12345678901231", EncodeTypes.GS1Code128);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "GS1Code128" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    [Fact]
    public void Execute_RecognizeGs1DataMatrix_ReturnsResults()
    {
        var imagePath = CreateTestBarcodeImage("recognize_gs1datamatrix.png", "(01)12345678901231",
            EncodeTypes.GS1DataMatrix);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "GS1DataMatrix" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.True(recognizeResult.Count > 0);
    }

    #endregion

    #region Barcode Info Properties

    [Fact]
    public void Execute_BarcodeInfoHasConfidenceValue()
    {
        var imagePath = CreateTestBarcodeImage("confidence_test.png", "ConfidenceTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.NotEmpty(recognizeResult.Barcodes);
        foreach (var barcode in recognizeResult.Barcodes) Assert.NotNull(barcode.Confidence);
    }

    [Fact]
    public void Execute_BarcodeInfoHasRegionValue()
    {
        var imagePath = CreateTestBarcodeImage("region_test.png", "RegionTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.NotEmpty(recognizeResult.Barcodes);
        foreach (var barcode in recognizeResult.Barcodes) Assert.NotNull(barcode.Region);
    }

    [Fact]
    public void Execute_ReturnsCorrectCodeText()
    {
        var expectedText = "EXPECTED123";
        var imagePath = CreateTestBarcodeImage("codetext_verify.png", expectedText, EncodeTypes.Code39Standard);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "Code39Standard" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.NotEmpty(recognizeResult.Barcodes);
        Assert.Contains(recognizeResult.Barcodes, b => b.CodeText == expectedText);
    }

    #endregion

    #region No Barcode Scenarios

    [Fact]
    public void Execute_WithEmptyImage_ReturnsZeroCount()
    {
        var imagePath = CreateEmptyImage("empty_image.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.Equal(0, recognizeResult.Count);
        Assert.Empty(recognizeResult.Barcodes);
    }

    [Fact]
    public void Execute_WithEmptyImage_ReturnsNoBarcodesMessage()
    {
        var imagePath = CreateEmptyImage("empty_image_msg.png");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.Contains("No barcodes found", recognizeResult.Message);
    }

    #endregion

    #region Decode Type Display

    [Fact]
    public void Execute_WithSpecificType_DisplaysCorrectDecodeType()
    {
        var imagePath = CreateTestBarcodeImage("decodetype_display.png", "DisplayTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "QR" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.Equal("QR", recognizeResult.DecodeType);
    }

    [Fact]
    public void Execute_WithAutoType_DisplaysAllAsDecodeType()
    {
        var imagePath = CreateTestBarcodeImage("decodetype_auto.png", "AutoTest", EncodeTypes.QR);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", imagePath },
            { "type", "auto" }
        });

        var result = _handler.Execute(context, parameters);

        var recognizeResult = Assert.IsType<RecognizeBarcodeResult>(result);
        Assert.Equal("All", recognizeResult.DecodeType);
    }

    #endregion
}
