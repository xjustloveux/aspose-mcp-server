using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.BarCode;

/// <summary>
///     Tool for recognizing barcodes from image files.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.BarCode.Recognize")]
[McpServerToolType]
public class BarcodeRecognizeTool
{
    /// <summary>
    ///     Handler registry for barcode recognition operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="BarcodeRecognizeTool" /> class.
    /// </summary>
    public BarcodeRecognizeTool()
    {
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.BarCode.Recognize");
    }

    /// <summary>
    ///     Executes a barcode recognition operation.
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: recognize.
    /// </param>
    /// <param name="path">Source image file path containing barcode(s).</param>
    /// <param name="type">Barcode type filter for recognition (default: auto, recognizes all types).</param>
    /// <returns>A barcode recognition result containing recognized barcode information.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "barcode_recognize",
        Title = "Barcode Recognizer",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = true,
        UseStructuredContent = true)]
    [Description(@"Recognize and decode barcodes from image files. Supports 1 operation: recognize.

Usage examples:
- Auto detect all barcodes: barcode_recognize(operation='recognize', path='image.png')
- QR code only: barcode_recognize(operation='recognize', path='image.png', type='QR')
- Code128 only: barcode_recognize(operation='recognize', path='image.png', type='Code128')
- EAN13 only: barcode_recognize(operation='recognize', path='image.png', type='EAN13')

Supported decode types: auto (all types), QR, Code128, Code39, EAN13, EAN8, UPCA, UPCE, DataMatrix, PDF417, Aztec, Code93, Codabar, ITF14, Interleaved2of5, GS1Code128, GS1DataMatrix
Supported image formats: PNG, JPEG, BMP, GIF, TIFF")]
    public object Execute(
        [Description(@"Operation to perform.
- 'recognize': Recognize barcodes from an image (required params: path)")]
        string operation,
        [Description("Source image file path containing barcode(s) (PNG, JPEG, BMP, GIF, TIFF)")]
        string? path = null,
        [Description(
            "Barcode type filter (default: auto). Options: auto, QR, Code128, Code39, EAN13, EAN8, UPCA, UPCE, DataMatrix, PDF417, Aztec, Code93, Codabar, ITF14")]
        string? type = null)
    {
        var parameters = BuildParameters(path, type);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = path,
            OutputPath = null
        };

        var result = handler.Execute(operationContext, parameters);

        return ResultHelper.FinalizeResult((dynamic)result, (string?)null, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The source image file path.</param>
    /// <param name="type">The barcode decode type filter.</param>
    /// <returns>OperationParameters configured for the recognition operation.</returns>
    private static OperationParameters BuildParameters(string? path, string? type)
    {
        var parameters = new OperationParameters();
        parameters.SetIfNotNull("path", path);
        parameters.SetIfNotNull("type", type);
        return parameters;
    }
}
