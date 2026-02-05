using Aspose.BarCode.Generation;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Base class for BarCode tool tests providing BarCode-specific functionality.
///     BarCode tools are stateless — generate/recognize operations do not use Sessions.
/// </summary>
public abstract class BarcodeTestBase : TestBase
{
    /// <summary>
    ///     Creates a barcode image file for testing.
    /// </summary>
    /// <param name="fileName">The output image file name.</param>
    /// <param name="text">The text to encode in the barcode.</param>
    /// <param name="encodeType">The barcode type (default: QR).</param>
    /// <returns>The full path to the created barcode image.</returns>
    protected string CreateBarcodeImage(string fileName, string text = "TestData", BaseEncodeType? encodeType = null)
    {
        var filePath = CreateTestFilePath(fileName);
        var generator = new BarcodeGenerator(encodeType ?? EncodeTypes.QR, text);
        generator.Save(filePath, BarCodeImageFormat.Png);
        return filePath;
    }

    /// <summary>
    ///     Checks if Aspose.BarCode is running in evaluation mode.
    /// </summary>
    protected new static bool IsEvaluationMode(AsposeLibraryType libraryType = AsposeLibraryType.BarCode)
    {
        return TestBase.IsEvaluationMode(libraryType);
    }
}
