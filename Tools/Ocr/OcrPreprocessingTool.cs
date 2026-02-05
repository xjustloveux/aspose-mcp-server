using System.ComponentModel;
using Aspose.OCR;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Ocr;

/// <summary>
///     Tool for preprocessing images to improve OCR recognition results.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Ocr.Preprocessing")]
[McpServerToolType]
public class OcrPreprocessingTool
{
    /// <summary>
    ///     Handler registry for OCR preprocessing operations.
    /// </summary>
    private readonly HandlerRegistry<AsposeOcr> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="OcrPreprocessingTool" /> class.
    /// </summary>
    public OcrPreprocessingTool()
    {
        _handlerRegistry =
            HandlerRegistry<AsposeOcr>.CreateFromNamespace("AsposeMcpServer.Handlers.Ocr.Preprocessing");
    }

    /// <summary>
    ///     Executes an OCR preprocessing operation (auto_skew, denoise, contrast, scale, invert, dewarp).
    /// </summary>
    /// <param name="operation">The preprocessing operation to perform.</param>
    /// <param name="path">Input image file path.</param>
    /// <param name="outputPath">Output file path for the preprocessed image.</param>
    /// <param name="scaleFactor">Scale factor for the scale operation (default: 2.0).</param>
    /// <returns>A preprocessing result containing operation details and output file info.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="InvalidOperationException">Thrown when preprocessing produces no output.</exception>
    [McpServerTool(
        Name = "ocr_preprocessing",
        Title = "OCR Image Preprocessing",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Preprocess images for better OCR results. Supports 6 operations: auto_skew, denoise, contrast, scale, invert, dewarp.

Usage examples:
- Auto-skew correction: ocr_preprocessing(operation='auto_skew', path='input.png', outputPath='output.png')
- Denoise image: ocr_preprocessing(operation='denoise', path='input.png', outputPath='output.png')
- Adjust contrast: ocr_preprocessing(operation='contrast', path='input.png', outputPath='output.png')
- Scale image: ocr_preprocessing(operation='scale', path='input.png', outputPath='output.png', scaleFactor=2.0)
- Invert colors: ocr_preprocessing(operation='invert', path='input.png', outputPath='output.png')
- Dewarp image: ocr_preprocessing(operation='dewarp', path='input.png', outputPath='output.png')

Supported image formats: PNG, JPG, BMP, TIFF, GIF
Note: OCR requires ONNX Runtime and is not supported on Linux ARM64.")]
    public object Execute(
        [Description(@"Preprocessing operation to perform.
- 'auto_skew': Automatically correct image tilt/rotation
- 'denoise': Remove noise and artifacts
- 'contrast': Enhance contrast for better text visibility
- 'scale': Enlarge or reduce image (use scaleFactor parameter)
- 'invert': Invert image colors (for white-on-black text)
- 'dewarp': Correct perspective distortion")]
        string operation,
        [Description("Input image file path (PNG, JPG, BMP, TIFF, GIF)")]
        string path,
        [Description("Output file path for the preprocessed image")]
        string outputPath,
        [Description("Scale factor for the scale operation (default: 2.0)")]
        double scaleFactor = 2.0)
    {
        var ocr = new AsposeOcr();

        var parameters = BuildParameters(path, outputPath, scaleFactor);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<AsposeOcr>
        {
            Document = ocr,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        return ResultHelper.FinalizeResult((dynamic)result, outputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The input image file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="scaleFactor">The scale factor for the scale operation.</param>
    /// <returns>OperationParameters configured for the preprocessing operation.</returns>
    private static OperationParameters BuildParameters(string path, string outputPath, double scaleFactor)
    {
        var parameters = new OperationParameters();
        parameters.Set("path", path);
        parameters.Set("outputPath", outputPath);
        parameters.Set("scaleFactor", scaleFactor);
        return parameters;
    }
}
