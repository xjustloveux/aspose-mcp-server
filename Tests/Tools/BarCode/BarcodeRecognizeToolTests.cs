using Aspose.BarCode.Generation;
using AsposeMcpServer.Results.BarCode.Generate;
using AsposeMcpServer.Results.BarCode.Recognize;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.BarCode;

namespace AsposeMcpServer.Tests.Tools.BarCode;

/// <summary>
///     Integration tests for <see cref="BarcodeRecognizeTool" />.
///     Focuses on operation routing, barcode recognition, and end-to-end workflows.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class BarcodeRecognizeToolTests : BarcodeTestBase
{
    private readonly BarcodeRecognizeTool _tool = new();

    #region End-to-End Workflow

    [SkippableFact]
    public void GenerateThenRecognize_EndToEnd_ShouldMatchText()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var generateTool = new BarcodeGenerateTool();
        var outputPath = CreateTestFilePath("e2e_barcode.png");
        var expectedText = "End2EndTest";

        var generateResult = generateTool.Execute("generate", expectedText, outputPath);
        var generateData = GetResultData<GenerateBarcodeResult>(generateResult);
        Assert.True(File.Exists(generateData.OutputPath));

        var recognizeResult = _tool.Execute("recognize", outputPath);
        var recognizeData = GetResultData<RecognizeBarcodeResult>(recognizeResult);

        Assert.True(recognizeData.Count > 0);
        Assert.Contains(recognizeData.Barcodes, b => b.CodeText == expectedText);
    }

    #endregion

    #region Recognize Operations

    [SkippableFact]
    public void Recognize_QrCode_ShouldRecognize()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_qr.png", "QR Test Data");

        var result = _tool.Execute("recognize", imagePath);

        Assert.NotNull(result);
        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.True(data.Count > 0);
        Assert.NotEmpty(data.Barcodes);
        Assert.Contains(data.Barcodes, b => b.CodeText == "QR Test Data");
    }

    [SkippableFact]
    public void Recognize_Code128_ShouldRecognize()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_code128.png", "Code128Data", EncodeTypes.Code128);

        var result = _tool.Execute("recognize", imagePath);

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Barcodes, b => b.CodeText == "Code128Data");
    }

    [SkippableFact]
    public void Recognize_WithTypeFilter_QR_ShouldRecognize()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_qr_filter.png", "FilteredQR");

        var result = _tool.Execute("recognize", imagePath, "QR");

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Barcodes, b => b.CodeText == "FilteredQR");
    }

    [SkippableFact]
    public void Recognize_WithAutoType_ShouldRecognize()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_auto.png", "AutoDetect");

        var result = _tool.Execute("recognize", imagePath, "auto");

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Barcodes, b => b.CodeText == "AutoDetect");
    }

    #endregion

    #region Result Validation

    [SkippableFact]
    public void Recognize_ResultHasCorrectSourcePath()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_source.png", "SourcePathTest");

        var result = _tool.Execute("recognize", imagePath);

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.Equal(imagePath, data.SourcePath);
    }

    [SkippableFact]
    public void Recognize_ResultHasPositiveCount()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_count.png", "CountTest");

        var result = _tool.Execute("recognize", imagePath);

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.True(data.Count > 0);
    }

    [SkippableFact]
    public void Recognize_BarcodesListNotEmpty()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_list.png", "ListTest");

        var result = _tool.Execute("recognize", imagePath);

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.NotEmpty(data.Barcodes);
    }

    [SkippableFact]
    public void Recognize_BarcodeInfoHasMatchingCodeText()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var encodedText = "MatchingTextTest";
        var imagePath = CreateBarcodeImage("recognize_codetext.png", encodedText);

        var result = _tool.Execute("recognize", imagePath);

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.Contains(data.Barcodes, b => b.CodeText == encodedText);
    }

    [SkippableFact]
    public void Recognize_BarcodeInfoHasMatchingCodeType()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_codetype.png", "CodeTypeTest", EncodeTypes.QR);

        var result = _tool.Execute("recognize", imagePath);

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.Contains(data.Barcodes, b => b.CodeType == "QR");
    }

    [SkippableFact]
    public void Recognize_MessageContainsFound()
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage("recognize_message.png", "MessageTest");

        var result = _tool.Execute("recognize", imagePath);

        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.Contains("Found", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Operation Routing

    [SkippableTheory]
    [InlineData("RECOGNIZE")]
    [InlineData("Recognize")]
    [InlineData("recognize")]
    public void Execute_CaseInsensitiveOperation_ShouldWork(string operation)
    {
        SkipInEvaluationMode(AsposeLibraryType.BarCode);
        var imagePath = CreateBarcodeImage($"case_{operation}.png", "CaseTest");

        var result = _tool.Execute(operation, imagePath);

        Assert.NotNull(result);
        var data = GetResultData<RecognizeBarcodeResult>(result);
        Assert.True(data.Count > 0);
    }

    [Fact]
    public void Execute_UnknownOperation_ShouldThrowArgumentException()
    {
        var imagePath = CreateBarcodeImage("unknown_op.png", "Unknown");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", imagePath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Recognize_MissingPath_ShouldThrowArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("recognize"));
    }

    [Fact]
    public void Recognize_NonExistentFile_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_barcode.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("recognize", fakePath));
    }

    #endregion
}
