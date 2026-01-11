using System.Drawing.Imaging;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for extracting images from PDF documents.
/// </summary>
public class ExtractPdfImageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts images from the specified page of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex, imageIndex, outputPath, outputDir
    /// </param>
    /// <returns>Success message with extraction details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var imageIndex = parameters.GetOptional<int?>("imageIndex");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var outputDir = parameters.GetOptional<string?>("outputDir");

        if (!string.IsNullOrEmpty(outputPath))
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
        if (!string.IsNullOrEmpty(outputDir))
            SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        var targetDir = outputDir ??
                        Path.GetDirectoryName(outputPath) ?? Path.GetDirectoryName(context.SourcePath) ?? ".";
        Directory.CreateDirectory(targetDir);

        var document = context.Document;
        var actualPageIndex = pageIndex < 1 ? 1 : pageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        var images = page.Resources?.Images;
        if (images == null || images.Count == 0)
            return Success($"No images found on page {pageIndex}.");

        if (imageIndex is > 0)
        {
            if (imageIndex.Value < 1 || imageIndex.Value > images.Count)
                throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

            var image = images[imageIndex.Value];
            var fileName = outputPath ?? Path.Combine(targetDir, $"page_{pageIndex}_image_{imageIndex.Value}.png");
            using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416
            image.Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
            return Success($"Extracted image {imageIndex.Value} from page {pageIndex} to: {fileName}");
        }

        var count = 0;
        for (var i = 1; i <= images.Count; i++)
        {
            var image = images[i];
            var fileName = Path.Combine(targetDir, $"page_{pageIndex}_image_{i}.png");
            using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416
            image.Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
            count++;
        }

        return Success($"Extracted {count} image(s) from page {pageIndex} to: {targetDir}");
    }
}
