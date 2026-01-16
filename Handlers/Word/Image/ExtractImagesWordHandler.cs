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
        var p = ExtractExtractImagesParameters(parameters);

        SecurityHelper.ValidateFilePath(p.OutputDir, "outputDir", true);

        Directory.CreateDirectory(p.OutputDir);

        var doc = context.Document;
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();

        if (shapes.Count == 0) return "No images found in document";

        if (p.ExtractImageIndex.HasValue &&
            (p.ExtractImageIndex.Value < 0 || p.ExtractImageIndex.Value >= shapes.Count))
            throw new ArgumentException(
                $"Image index {p.ExtractImageIndex.Value} is out of range (document has {shapes.Count} images)");

        List<string> extractedFiles = [];

        var startIndex = p.ExtractImageIndex ?? 0;
        var endIndex = p.ExtractImageIndex.HasValue ? p.ExtractImageIndex.Value + 1 : shapes.Count;

        for (var i = startIndex; i < endIndex; i++)
        {
            var shape = shapes[i];
            var imageData = shape.ImageData;

            var extension = FileFormatUtil.ImageTypeToExtension(imageData.ImageType);
            if (string.IsNullOrEmpty(extension) || extension == ".")
                extension = ".img";
            if (extension.StartsWith('.'))
                extension = extension.Substring(1);

            var safePrefix = SecurityHelper.SanitizeFileName(p.Prefix);
            var filename = $"{safePrefix}_{i + 1:D3}.{extension}";
            var outputFilePath = Path.Combine(p.OutputDir, filename);

            using (var stream = IOFile.Create(outputFilePath))
            {
                imageData.Save(stream);
            }

            extractedFiles.Add(outputFilePath);
        }

        if (p.ExtractImageIndex.HasValue)
            return $"Successfully extracted image #{p.ExtractImageIndex.Value} to: {p.OutputDir}\n" +
                   $"File: {Path.GetFileName(extractedFiles[0])}";

        return $"Successfully extracted {shapes.Count} images to: {p.OutputDir}\n" +
               $"File list:\n" + string.Join("\n",
                   extractedFiles.Select(f => $"  - {Path.GetFileName(f)}"));
    }

    private static ExtractImagesParameters ExtractExtractImagesParameters(OperationParameters parameters)
    {
        return new ExtractImagesParameters(
            parameters.GetRequired<string>("outputDir"),
            parameters.GetOptional("prefix", "image"),
            parameters.GetOptional<int?>("extractImageIndex"));
    }

    private record ExtractImagesParameters(
        string OutputDir,
        string Prefix,
        int? ExtractImageIndex);
}
