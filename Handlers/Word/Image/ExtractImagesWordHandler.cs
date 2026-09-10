using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for extracting images from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractExtractImagesParameters(parameters);

        SecurityHelper.ValidateFilePath(p.OutputDir, "outputDir", true);
        var resolvedOutputDir = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputDir,
            context.ServerConfig?.AllowedBasePaths ?? [], "outputDir");

        Directory.CreateDirectory(resolvedOutputDir);

        var doc = context.Document;
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();

        if (shapes.Count == 0) return new SuccessResult { Message = "No images found in document" };

        if (p.ExtractImageIndex.HasValue &&
            (p.ExtractImageIndex.Value < 0 || p.ExtractImageIndex.Value >= shapes.Count))
            throw new ArgumentException(
                $"Image index {p.ExtractImageIndex.Value} is out of range (document has {shapes.Count} images)");

        List<string> extractedFiles = [];

        var startIndex = p.ExtractImageIndex ?? 0;
        var endIndex = p.ExtractImageIndex.HasValue ? p.ExtractImageIndex.Value + 1 : shapes.Count;

        // A document can hold one image per shape, and each becomes a file (R2-R02).
        RenderBudget.EnsureOutputCount(endIndex - startIndex, "image files");

        // Staged as a batch and published once: an extraction refused part-way used to leave
        // whichever images it had already produced on top of the caller's files (R5-R01).
        using var batch = new BoundedFileBatch(RenderBudget.MaxOutputBytes, "extracted images",
            context.Recovery, context.ServerConfig?.AllowedBasePaths ?? []);

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
            var outputFilePath = Path.Combine(resolvedOutputDir, filename);
            // H11: resolve symlinks immediately before the File.Create sink (bug 20260415-symlink-toctou-sweep).
            outputFilePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputFilePath,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(outputFilePath));

            // Measured after the write, an image above the limit was already on disk when the
            // request was refused (R3-R07).
            batch.Stage(outputFilePath, imageData.Save);

            extractedFiles.Add(outputFilePath);
        }

        batch.Publish();

        if (p.ExtractImageIndex.HasValue)
            return new SuccessResult
            {
                Message = $"Successfully extracted image #{p.ExtractImageIndex.Value} to: {resolvedOutputDir}\n" +
                          $"File: {Path.GetFileName(extractedFiles[0])}"
            };

        return new SuccessResult
        {
            Message = $"Successfully extracted {shapes.Count} images to: {resolvedOutputDir}\n" +
                      $"File list:\n" + string.Join("\n",
                          extractedFiles.Select(f => $"  - {Path.GetFileName(f)}"))
        };
    }

    private static ExtractImagesParameters ExtractExtractImagesParameters(OperationParameters parameters)
    {
        return new ExtractImagesParameters(
            parameters.GetRequired<string>("outputDir"),
            parameters.GetOptional("prefix", "image"),
            parameters.GetOptional<int?>("extractImageIndex"));
    }

    private sealed record ExtractImagesParameters(
        string OutputDir,
        string Prefix,
        int? ExtractImageIndex);
}
