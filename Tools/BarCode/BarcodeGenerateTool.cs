using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.BarCode;

/// <summary>
///     Tool for generating barcode images in various formats.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.BarCode.Generate")]
[McpServerToolType]
public class BarcodeGenerateTool
{
    /// <summary>
    ///     Handler registry for barcode generation operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="BarcodeGenerateTool" /> class.
    /// </summary>
    public BarcodeGenerateTool()
    {
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.BarCode.Generate");
    }

    /// <summary>
    ///     Executes a barcode generation operation.
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: generate.
    /// </param>
    /// <param name="text">The text or data to encode in the barcode.</param>
    /// <param name="outputPath">Output image file path (format determined by extension: .png, .jpg, .bmp, .gif, .tiff, .svg).</param>
    /// <param name="type">
    ///     Barcode type: QR, Code128, Code39, EAN13, EAN8, UPCA, UPCE, DataMatrix, PDF417, Aztec, Code93,
    ///     Codabar, ITF14. Default: QR.
    /// </param>
    /// <param name="width">Barcode X-dimension width in pixels.</param>
    /// <param name="height">Barcode bar height in pixels.</param>
    /// <param name="foreColor">Foreground (bar) color as color name or hex (#RRGGBB). Default: Black.</param>
    /// <param name="backColor">Background color as color name or hex (#RRGGBB). Default: White.</param>
    /// <returns>A barcode generation result containing file and barcode information.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "barcode_generate",
        Title = "Barcode Generator",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Generate barcode images in various formats. Supports 1 operation: generate.

Usage examples:
- QR Code: barcode_generate(operation='generate', text='Hello World', outputPath='qr.png')
- Code128: barcode_generate(operation='generate', text='12345', outputPath='code128.png', type='Code128')
- EAN13: barcode_generate(operation='generate', text='5901234123457', outputPath='ean.png', type='EAN13')
- Custom colors: barcode_generate(operation='generate', text='Test', outputPath='custom.png', foreColor='#FF0000', backColor='#FFFF00')

Supported barcode types: QR, Code128, Code39, EAN13, EAN8, UPCA, UPCE, DataMatrix, PDF417, Aztec, Code93, Codabar, ITF14, Interleaved2of5, GS1Code128, GS1DataMatrix
Supported image formats: PNG, JPEG, BMP, GIF, TIFF, SVG, EMF")]
    public object Execute(
        [Description(@"Operation to perform.
- 'generate': Generate a barcode image (required params: text, outputPath)")]
        string operation,
        [Description("The text or data to encode in the barcode")]
        string? text = null,
        [Description(
            "Output image file path (format determined by extension: .png, .jpg, .bmp, .gif, .tiff, .svg, .emf)")]
        string? outputPath = null,
        [Description(
            "Barcode type (default: QR). Options: QR, Code128, Code39, EAN13, EAN8, UPCA, UPCE, DataMatrix, PDF417, Aztec, Code93, Codabar, ITF14")]
        string? type = null,
        [Description("Barcode X-dimension width in pixels")]
        int? width = null,
        [Description("Barcode bar height in pixels")]
        int? height = null,
        [Description("Foreground (bar) color as color name or hex (#RRGGBB). Default: Black")]
        string? foreColor = null,
        [Description("Background color as color name or hex (#RRGGBB). Default: White")]
        string? backColor = null)
    {
        var parameters = BuildParameters(text, outputPath, type, width, height, foreColor, backColor);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = null,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        return ResultHelper.FinalizeResult((dynamic)result, outputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="text">The text to encode.</param>
    /// <param name="outputPath">The output image file path.</param>
    /// <param name="type">The barcode type.</param>
    /// <param name="width">The barcode width.</param>
    /// <param name="height">The barcode height.</param>
    /// <param name="foreColor">The foreground color.</param>
    /// <param name="backColor">The background color.</param>
    /// <returns>OperationParameters configured for the generation operation.</returns>
    private static OperationParameters BuildParameters(string? text, string? outputPath, string? type,
        int? width, int? height, string? foreColor, string? backColor)
    {
        var parameters = new OperationParameters();
        parameters.SetIfNotNull("text", text);
        parameters.SetIfNotNull("outputPath", outputPath);
        parameters.SetIfNotNull("type", type);
        parameters.SetIfHasValue("width", width);
        parameters.SetIfHasValue("height", height);
        parameters.SetIfNotNull("foreColor", foreColor);
        parameters.SetIfNotNull("backColor", backColor);
        return parameters;
    }
}
