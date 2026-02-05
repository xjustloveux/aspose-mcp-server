using Aspose.BarCode.BarCodeRecognition;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.BarCode.Recognize;

namespace AsposeMcpServer.Handlers.BarCode.Recognize;

/// <summary>
///     Handler for recognizing barcodes from image files.
/// </summary>
[ResultType(typeof(RecognizeBarcodeResult))]
public class RecognizeBarcodeHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "recognize";

    /// <summary>
    ///     Recognizes barcodes from the specified image file.
    /// </summary>
    /// <param name="context">The operation context (not used for barcode operations).</param>
    /// <param name="parameters">
    ///     Required: path (source image file).
    ///     Optional: type (barcode decode type filter, default: AllSupportedTypes).
    /// </param>
    /// <returns>A <see cref="RecognizeBarcodeResult" /> containing recognition details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or decode type is unsupported.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");

        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Input file not found: {path}");

        var typeStr = parameters.GetOptional("type", "AllSupportedTypes");
        var decodeType = GetDecodeType(typeStr);

        var barcodes = new List<BarcodeInfo>();

        using (var reader = new BarCodeReader(path, decodeType))
        {
            foreach (var result in reader.ReadBarCodes())
                barcodes.Add(new BarcodeInfo
                {
                    CodeText = result.CodeText,
                    CodeType = result.CodeTypeName,
                    Confidence = result.Confidence.ToString(),
                    Region = result.Region?.Rectangle.ToString()
                });
        }

        var displayType = typeStr.Equals("AllSupportedTypes", StringComparison.OrdinalIgnoreCase) ||
                          typeStr.Equals("all", StringComparison.OrdinalIgnoreCase) ||
                          typeStr.Equals("auto", StringComparison.OrdinalIgnoreCase)
            ? "All"
            : typeStr;

        var message = barcodes.Count == 0
            ? $"No barcodes found in: {path}"
            : $"Found {barcodes.Count} barcode(s) in: {path}";

        return new RecognizeBarcodeResult
        {
            SourcePath = path,
            Barcodes = barcodes,
            Count = barcodes.Count,
            DecodeType = displayType,
            Message = message
        };
    }

    /// <summary>
    ///     Gets the Aspose.BarCode decode type from a string name.
    /// </summary>
    /// <param name="type">The decode type name (case-insensitive).</param>
    /// <returns>The corresponding <see cref="BaseDecodeType" />.</returns>
    /// <exception cref="ArgumentException">Thrown when the decode type is not supported.</exception>
    private static BaseDecodeType GetDecodeType(string type)
    {
        return type.ToUpperInvariant() switch
        {
            "AUTO" or "ALL" or "ALLSUPPORTEDTYPES" => DecodeType.AllSupportedTypes,
            "QR" or "QRCODE" => DecodeType.QR,
            "CODE128" => DecodeType.Code128,
            "CODE39" or "CODE39STANDARD" => DecodeType.Code39Standard,
            "EAN13" => DecodeType.EAN13,
            "EAN8" => DecodeType.EAN8,
            "UPCA" => DecodeType.UPCA,
            "UPCE" => DecodeType.UPCE,
            "DATAMATRIX" => DecodeType.DataMatrix,
            "PDF417" => DecodeType.Pdf417,
            "AZTEC" => DecodeType.Aztec,
            "CODE93" or "CODE93STANDARD" => DecodeType.Code93Standard,
            "CODABAR" => DecodeType.Codabar,
            "ITF14" => DecodeType.ITF14,
            "INTERLEAVED2OF5" => DecodeType.Interleaved2of5,
            "GS1CODE128" => DecodeType.GS1Code128,
            "GS1DATAMATRIX" => DecodeType.GS1DataMatrix,
            _ => throw new ArgumentException(
                $"Unsupported decode type: {type}. Supported types: auto, QR, Code128, Code39, EAN13, EAN8, " +
                "UPCA, UPCE, DataMatrix, PDF417, Aztec, Code93, Codabar, ITF14, Interleaved2of5, " +
                "GS1Code128, GS1DataMatrix")
        };
    }
}
