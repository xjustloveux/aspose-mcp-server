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
}
