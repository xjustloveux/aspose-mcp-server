using Aspose.Pdf;
using Aspose.Pdf.Optimization;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for compressing a PDF document.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var compressImages = parameters.GetOptional("compressImages", true);
        var compressFonts = parameters.GetOptional("compressFonts", true);
        var removeUnusedObjects = parameters.GetOptional("removeUnusedObjects", true);

        var document = context.Document;
        var optimizationOptions = new OptimizationOptions();

        if (compressImages)
        {
            optimizationOptions.ImageCompressionOptions.CompressImages = true;
            optimizationOptions.ImageCompressionOptions.ImageQuality = 75;
        }

        if (compressFonts)
            optimizationOptions.SubsetFonts = true;

        if (removeUnusedObjects)
        {
            optimizationOptions.LinkDuplcateStreams = true;
            optimizationOptions.RemoveUnusedObjects = true;
            optimizationOptions.AllowReusePageContent = true;
        }

        document.OptimizeResources(optimizationOptions);

        MarkModified(context);

        return Success("PDF compressed.");
    }
}
