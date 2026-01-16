using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using IOFile = System.IO.File;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for extracting images from Word documents.
/// </summary>
public class ExtractImagesWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts images from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputDir
    ///     Optional: prefix, extractImageIndex
    /// </param>
    /// <returns>Success message with extraction details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var outputDir = parameters.GetRequired<string>("outputDir");
        var prefix = parameters.GetOptional("prefix", "image");
        var extractImageIndex = parameters.GetOptional<int?>("extractImageIndex");

        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        Directory.CreateDirectory(outputDir);

        var doc = context.Document;
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();

        if (shapes.Count == 0) return "No images found in document";

        // Validate extractImageIndex if provided
        if (extractImageIndex.HasValue &&
            (extractImageIndex.Value < 0 || extractImageIndex.Value >= shapes.Count))
            throw new ArgumentException(
                $"Image index {extractImageIndex.Value} is out of range (document has {shapes.Count} images)");

        List<string> extractedFiles = [];

        // Determine which images to extract
        var startIndex = extractImageIndex ?? 0;
        var endIndex = extractImageIndex.HasValue ? extractImageIndex.Value + 1 : shapes.Count;

        for (var i = startIndex; i < endIndex; i++)
        {
            var shape = shapes[i];
            var imageData = shape.ImageData;

            // Use FileFormatUtil for reliable image type detection
            var extension = FileFormatUtil.ImageTypeToExtension(imageData.ImageType);
            if (string.IsNullOrEmpty(extension) || extension == ".")
                extension = ".img";
            // Remove leading dot if present for consistent filename handling
            if (extension.StartsWith('.'))
                extension = extension.Substring(1);

            var safePrefix = SecurityHelper.SanitizeFileName(prefix);
            var filename = $"{safePrefix}_{i + 1:D3}.{extension}";
            var outputFilePath = Path.Combine(outputDir, filename);

            using (var stream = IOFile.Create(outputFilePath))
            {
                imageData.Save(stream);
            }

            extractedFiles.Add(outputFilePath);
        }

        if (extractImageIndex.HasValue)
            return $"Successfully extracted image #{extractImageIndex.Value} to: {outputDir}\n" +
                   $"File: {Path.GetFileName(extractedFiles[0])}";

        return $"Successfully extracted {shapes.Count} images to: {outputDir}\n" +
               $"File list:\n" + string.Join("\n",
                   extractedFiles.Select(f => $"  - {Path.GetFileName(f)}"));
    }
}
