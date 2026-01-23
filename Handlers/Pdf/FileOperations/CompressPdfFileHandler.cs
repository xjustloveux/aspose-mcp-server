using Aspose.Pdf;
using Aspose.Pdf.Optimization;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for compressing a PDF document.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CompressPdfFileHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "compress";

    /// <summary>
    ///     Compresses a PDF document by optimizing images, fonts, and removing unused objects.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: compressImages (default: true), compressFonts (default: true), removeUnusedObjects (default: true)
    /// </param>
    /// <returns>Success message with compression statistics.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var compressParams = ExtractCompressParameters(parameters);

        var document = context.Document;
        var optimizationOptions = new OptimizationOptions();

        if (compressParams.CompressImages)
        {
            optimizationOptions.ImageCompressionOptions.CompressImages = true;
            optimizationOptions.ImageCompressionOptions.ImageQuality = 75;
        }

        if (compressParams.CompressFonts)
            optimizationOptions.SubsetFonts = true;

        if (compressParams.RemoveUnusedObjects)
        {
            optimizationOptions.LinkDuplcateStreams = true;
            optimizationOptions.RemoveUnusedObjects = true;
            optimizationOptions.AllowReusePageContent = true;
        }

        document.OptimizeResources(optimizationOptions);

        MarkModified(context);

        return new SuccessResult { Message = "PDF compressed." };
    }

    /// <summary>
    ///     Extracts compress parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted compress parameters.</returns>
    private static CompressParameters ExtractCompressParameters(OperationParameters parameters)
    {
        return new CompressParameters(
            parameters.GetOptional("compressImages", true),
            parameters.GetOptional("compressFonts", true),
            parameters.GetOptional("removeUnusedObjects", true)
        );
    }

    /// <summary>
    ///     Record to hold compress parameters.
    /// </summary>
    private sealed record CompressParameters(bool CompressImages, bool CompressFonts, bool RemoveUnusedObjects);
}
