using System.Drawing;
using Aspose.BarCode.Generation;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.BarCode.Generate;
using DrawingColor = Aspose.Drawing.Color;

namespace AsposeMcpServer.Handlers.BarCode.Generate;

/// <summary>
///     Handler for generating barcode images.
/// </summary>
[ResultType(typeof(GenerateBarcodeResult))]
public class GenerateBarcodeHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "generate";

    /// <summary>
    ///     Generates a barcode image and saves it to the specified output path.
    /// </summary>
    /// <param name="context">The operation context (not used for barcode operations).</param>
    /// <param name="parameters">
    ///     Required: text (content to encode), outputPath (destination image file).
    ///     Optional: type (barcode type), width, height, foreColor, backColor.
    /// </param>
    /// <returns>A <see cref="GenerateBarcodeResult" /> containing generation details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or barcode type is unsupported.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var text = parameters.GetRequired<string>("text");
        var outputPath = parameters.GetRequired<string>("outputPath");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var typeStr = parameters.GetOptional("type", "QR");
        var encodeType = GetEncodeType(typeStr);

        var generator = new BarcodeGenerator(encodeType, text);

        var width = parameters.GetOptional<int?>("width");
        var height = parameters.GetOptional<int?>("height");
        if (width.HasValue)
            generator.Parameters.Barcode.XDimension.Pixels = width.Value;
        if (height.HasValue)
            generator.Parameters.Barcode.BarHeight.Pixels = height.Value;

        var foreColor = parameters.GetOptional<string>("foreColor");
        var backColor = parameters.GetOptional<string>("backColor");
        if (!string.IsNullOrEmpty(foreColor))
            generator.Parameters.Barcode.BarColor = ToAsposeColor(ColorHelper.ParseColor(foreColor, true));
        if (!string.IsNullOrEmpty(backColor))
            generator.Parameters.BackColor = ToAsposeColor(ColorHelper.ParseColor(backColor, true));

        var imageFormat = GetImageFormat(outputPath);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        generator.Save(outputPath, imageFormat);

        var formatName = Path.GetExtension(outputPath).TrimStart('.').ToUpperInvariant();

        return new GenerateBarcodeResult
        {
            OutputPath = outputPath,
            BarcodeType = typeStr.ToUpperInvariant(),
            EncodedText = text,
            ImageFormat = formatName,
            FileSize = File.Exists(outputPath) ? new FileInfo(outputPath).Length : null,
            Message = $"Barcode ({typeStr.ToUpperInvariant()}) generated successfully: {outputPath}"
        };
    }

    /// <summary>
    ///     Gets the Aspose.BarCode encode type from a string name.
    /// </summary>
    /// <param name="type">The barcode type name (case-insensitive).</param>
    /// <returns>The corresponding <see cref="BaseEncodeType" />.</returns>
    /// <exception cref="ArgumentException">Thrown when the barcode type is not supported.</exception>
    private static BaseEncodeType GetEncodeType(string type)
    {
        return type.ToUpperInvariant() switch
        {
            "QR" or "QRCODE" => EncodeTypes.QR,
            "CODE128" => EncodeTypes.Code128,
            "CODE39" or "CODE39STANDARD" => EncodeTypes.Code39Standard,
            "CODE39EXTENDED" => EncodeTypes.Code39Extended,
            "EAN13" => EncodeTypes.EAN13,
            "EAN8" => EncodeTypes.EAN8,
            "UPCA" => EncodeTypes.UPCA,
            "UPCE" => EncodeTypes.UPCE,
            "DATAMATRIX" => EncodeTypes.DataMatrix,
            "PDF417" => EncodeTypes.Pdf417,
            "AZTEC" => EncodeTypes.Aztec,
            "CODE93STANDARD" or "CODE93" => EncodeTypes.Code93Standard,
            "CODE93EXTENDED" => EncodeTypes.Code93Extended,
            "CODABAR" => EncodeTypes.Codabar,
            "ITF14" => EncodeTypes.ITF14,
            "INTERLEAVED2OF5" => EncodeTypes.Interleaved2of5,
            "GS1CODE128" => EncodeTypes.GS1Code128,
            "GS1DATAMATRIX" => EncodeTypes.GS1DataMatrix,
            _ => throw new ArgumentException(
                $"Unsupported barcode type: {type}. Supported types: QR, Code128, Code39, EAN13, EAN8, " +
                "UPCA, UPCE, DataMatrix, PDF417, Aztec, Code93, Codabar, ITF14, Interleaved2of5, " +
                "GS1Code128, GS1DataMatrix")
        };
    }

    /// <summary>
    ///     Gets the barcode image format from the output file extension.
    /// </summary>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>The corresponding <see cref="BarCodeImageFormat" />.</returns>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    private static BarCodeImageFormat GetImageFormat(string outputPath)
    {
        var ext = Path.GetExtension(outputPath).ToLowerInvariant();
        return ext switch
        {
            ".png" => BarCodeImageFormat.Png,
            ".jpg" or ".jpeg" => BarCodeImageFormat.Jpeg,
            ".bmp" => BarCodeImageFormat.Bmp,
            ".gif" => BarCodeImageFormat.Gif,
            ".tiff" or ".tif" => BarCodeImageFormat.Tiff,
            ".svg" => BarCodeImageFormat.Svg,
            ".emf" => BarCodeImageFormat.Emf,
            _ => throw new ArgumentException(
                $"Unsupported image format: {ext}. Supported formats: png, jpg, jpeg, bmp, gif, tiff, tif, svg, emf")
        };
    }

    /// <summary>
    ///     Converts a <see cref="System.Drawing.Color" /> to an <see cref="DrawingColor" /> (Aspose.Drawing.Color).
    /// </summary>
    /// <param name="color">The System.Drawing.Color to convert.</param>
    /// <returns>The equivalent Aspose.Drawing.Color.</returns>
    private static DrawingColor ToAsposeColor(Color color)
    {
        return DrawingColor.FromArgb(color.A, color.R, color.G, color.B);
    }
}
