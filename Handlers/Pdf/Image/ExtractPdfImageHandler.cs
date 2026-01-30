using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using ImageFormat = Aspose.Pdf.Drawing.ImageFormat;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for extracting images from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractExtractParameters(parameters);

        if (!string.IsNullOrEmpty(p.OutputPath))
            SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);
        if (!string.IsNullOrEmpty(p.OutputDir))
            SecurityHelper.ValidateFilePath(p.OutputDir, "outputDir", true);

        var targetDir = p.OutputDir ??
                        Path.GetDirectoryName(p.OutputPath) ?? Path.GetDirectoryName(context.SourcePath) ?? ".";
        Directory.CreateDirectory(targetDir);

        var document = context.Document;
        var actualPageIndex = p.PageIndex < 1 ? 1 : p.PageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        var images = page.Resources?.Images;
        if (images == null || images.Count == 0)
            return new SuccessResult { Message = $"No images found on page {p.PageIndex}." };

        if (p.ImageIndex is > 0)
        {
            if (p.ImageIndex.Value < 1 || p.ImageIndex.Value > images.Count)
                throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

            var image = images[p.ImageIndex.Value];
            var fileName = p.OutputPath ??
                           Path.Combine(targetDir, $"page_{p.PageIndex}_image_{p.ImageIndex.Value}.png");
            using var imageStream = new FileStream(fileName, FileMode.Create);
            image.Save(imageStream, ImageFormat.Png);
            return new SuccessResult
                { Message = $"Extracted image {p.ImageIndex.Value} from page {p.PageIndex} to: {fileName}" };
        }

        var count = 0;
        for (var i = 1; i <= images.Count; i++)
        {
            var image = images[i];
            var fileName = Path.Combine(targetDir, $"page_{p.PageIndex}_image_{i}.png");
            using var imageStream = new FileStream(fileName, FileMode.Create);
            image.Save(imageStream, ImageFormat.Png);
            count++;
        }

        return new SuccessResult { Message = $"Extracted {count} image(s) from page {p.PageIndex} to: {targetDir}" };
    }

    /// <summary>
    ///     Extracts extract parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static ExtractParameters ExtractExtractParameters(OperationParameters parameters)
    {
        return new ExtractParameters(
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional<int?>("imageIndex"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("outputDir"));
    }

    /// <summary>
    ///     Parameters for extracting images.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="ImageIndex">The optional 1-based image index.</param>
    /// <param name="OutputPath">The optional output file path.</param>
    /// <param name="OutputDir">The optional output directory.</param>
    private sealed record ExtractParameters(int PageIndex, int? ImageIndex, string? OutputPath, string? OutputDir);
}
