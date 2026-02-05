using Aspose.Words;
using Aspose.Words.Saving;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Word.Render;

namespace AsposeMcpServer.Handlers.Word.Render;

/// <summary>
///     Handler for rendering a thumbnail of the first page of a Word document.
/// </summary>
[ResultType(typeof(RenderResult))]
public class RenderThumbnailWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "render_thumbnail";

    /// <summary>
    ///     Renders a thumbnail of the first page of a Word document.
    ///     The thumbnail is a lower-resolution image suitable for previews.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (source file path), outputPath (output image file path)
    ///     Optional: format (png, jpeg; default: png), scale (default: 0.25)
    /// </param>
    /// <returns>Render result with output file path.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or scale is invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractThumbnailParameters(parameters);

        SecurityHelper.ValidateFilePath(p.Path, allowAbsolutePaths: true);
        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        if (p.Scale <= 0 || p.Scale > 1)
            throw new ArgumentException("scale must be between 0 (exclusive) and 1 (inclusive)");

        var doc = new Document(p.Path);

        var saveFormat = p.Format.ToLowerInvariant() switch
        {
            "png" => SaveFormat.Png,
            "jpeg" or "jpg" => SaveFormat.Jpeg,
            _ => throw new ArgumentException(
                $"Unknown thumbnail format: {p.Format}. Supported: png, jpeg")
        };

        var options = new ImageSaveOptions(saveFormat)
        {
            Resolution = (float)(96 * p.Scale),
            PageSet = new PageSet(0)
        };

        var outputDir = Path.GetDirectoryName(p.OutputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        doc.Save(p.OutputPath, options);

        return new RenderResult
        {
            Message = $"Thumbnail rendered at {p.Scale:P0} scale in {p.Format.ToUpperInvariant()} format.",
            OutputPaths = [p.OutputPath],
            Format = p.Format
        };
    }

    /// <summary>
    ///     Extracts thumbnail parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted thumbnail parameters.</returns>
    private static ThumbnailParameters ExtractThumbnailParameters(OperationParameters parameters)
    {
        return new ThumbnailParameters(
            parameters.GetRequired<string>("path"),
            parameters.GetRequired<string>("outputPath"),
            parameters.GetOptional("format", "png"),
            parameters.GetOptional("scale", 0.25)
        );
    }

    /// <summary>
    ///     Parameters for the render_thumbnail operation.
    /// </summary>
    /// <param name="Path">The source document file path.</param>
    /// <param name="OutputPath">The output image file path.</param>
    /// <param name="Format">The output image format.</param>
    /// <param name="Scale">The scale factor (0-1).</param>
    private sealed record ThumbnailParameters(
        string Path,
        string OutputPath,
        string Format,
        double Scale);
}
