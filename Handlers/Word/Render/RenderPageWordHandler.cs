using Aspose.Words;
using Aspose.Words.Saving;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Word.Render;

namespace AsposeMcpServer.Handlers.Word.Render;

/// <summary>
///     Handler for rendering specific pages of a Word document to images.
/// </summary>
[ResultType(typeof(RenderResult))]
public class RenderPageWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "render_page";

    /// <summary>
    ///     Renders a specific page (or all pages) of a Word document to image file(s).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (source file path), outputPath (output file/directory path)
    ///     Optional: pageIndex (1-based, default: all pages), format (png, jpeg, bmp, tiff, svg; default: png), dpi (default:
    ///     150)
    /// </param>
    /// <returns>Render result with output file paths.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractRenderParameters(parameters);

        SecurityHelper.ValidateFilePath(p.Path, allowAbsolutePaths: true);
        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        var doc = new Document(p.Path);
        var saveFormat = ResolveSaveFormat(p.Format);
        var outputPaths = new List<string>();

        if (p.PageIndex.HasValue)
        {
            if (p.PageIndex.Value < 1 || p.PageIndex.Value > doc.PageCount)
                throw new ArgumentException(
                    $"pageIndex {p.PageIndex.Value} is out of range (1-{doc.PageCount})");

            var options = CreateImageSaveOptions(saveFormat, p.Dpi, p.PageIndex.Value - 1);

            var outputDir = Path.GetDirectoryName(p.OutputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            doc.Save(p.OutputPath, options);
            outputPaths.Add(p.OutputPath);
        }
        else
        {
            var outputDir = Path.GetDirectoryName(p.OutputPath);
            var baseName = Path.GetFileNameWithoutExtension(p.OutputPath);
            var ext = Path.GetExtension(p.OutputPath);
            if (string.IsNullOrEmpty(ext)) ext = $".{p.Format}";

            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            for (var i = 0; i < doc.PageCount; i++)
            {
                var options = CreateImageSaveOptions(saveFormat, p.Dpi, i);
                var pagePath = Path.Combine(outputDir ?? ".",
                    $"{baseName}_page_{i + 1}{ext}");
                doc.Save(pagePath, options);
                outputPaths.Add(pagePath);
            }
        }

        var message = p.PageIndex.HasValue
            ? $"Page {p.PageIndex.Value} rendered to {p.Format.ToUpperInvariant()} format."
            : $"{outputPaths.Count} page(s) rendered to {p.Format.ToUpperInvariant()} format.";

        return new RenderResult
        {
            Message = message,
            OutputPaths = outputPaths,
            Format = p.Format
        };
    }

    /// <summary>
    ///     Resolves the save format from a format string.
    /// </summary>
    /// <param name="format">The format string.</param>
    /// <returns>The corresponding SaveFormat value.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is unknown.</exception>
    private static SaveFormat ResolveSaveFormat(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "png" => SaveFormat.Png,
            "jpeg" or "jpg" => SaveFormat.Jpeg,
            "bmp" => SaveFormat.Bmp,
            "tiff" or "tif" => SaveFormat.Tiff,
            "svg" => SaveFormat.Svg,
            _ => throw new ArgumentException(
                $"Unknown render format: {format}. Supported: png, jpeg, bmp, tiff, svg")
        };
    }

    /// <summary>
    ///     Creates ImageSaveOptions for a specific page.
    /// </summary>
    /// <param name="saveFormat">The save format.</param>
    /// <param name="dpi">The rendering DPI.</param>
    /// <param name="pageIndex">The 0-based page index.</param>
    /// <returns>Configured ImageSaveOptions.</returns>
    private static ImageSaveOptions CreateImageSaveOptions(SaveFormat saveFormat, int dpi, int pageIndex)
    {
        return new ImageSaveOptions(saveFormat)
        {
            Resolution = dpi,
            PageSet = new PageSet(pageIndex)
        };
    }

    /// <summary>
    ///     Extracts render parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted render parameters.</returns>
    private static RenderParameters ExtractRenderParameters(OperationParameters parameters)
    {
        return new RenderParameters(
            parameters.GetRequired<string>("path"),
            parameters.GetRequired<string>("outputPath"),
            parameters.GetOptional<int?>("pageIndex"),
            parameters.GetOptional("format", "png"),
            parameters.GetOptional("dpi", 150)
        );
    }

    /// <summary>
    ///     Parameters for the render_page operation.
    /// </summary>
    /// <param name="Path">The source document file path.</param>
    /// <param name="OutputPath">The output file/directory path.</param>
    /// <param name="PageIndex">The 1-based page index (null for all pages).</param>
    /// <param name="Format">The output image format.</param>
    /// <param name="Dpi">The rendering DPI.</param>
    private sealed record RenderParameters(
        string Path,
        string OutputPath,
        int? PageIndex,
        string Format,
        int Dpi);
}
