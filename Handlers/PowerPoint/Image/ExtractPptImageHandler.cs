using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for extracting embedded images from PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ExtractPptImageHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts embedded images from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: (none, uses context.SourcePath)
    ///     Optional: outputDir, format, skipDuplicates
    /// </param>
    /// <returns>Success message with extraction details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var path = ValidateSourcePath(context.SourcePath);
        var extractParams = ExtractImageParameters(parameters, path);

        Directory.CreateDirectory(extractParams.OutputDir);

        var (count, skippedCount) = ExtractAllImages(context.Document, extractParams);

        return new SuccessResult { Message = BuildResultMessage(count, skippedCount, extractParams) };
    }

    /// <summary>
    ///     Validates the source path.
    /// </summary>
    /// <param name="path">The source path to validate.</param>
    /// <returns>The validated path.</returns>
    private static string ValidateSourcePath(string? path)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for extract operation");
        SecurityHelper.ValidateFilePath(path, nameof(path), true);
        return path;
    }

    /// <summary>
    ///     Extracts image parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <param name="path">The source file path.</param>
    /// <returns>The extraction parameters.</returns>
    private static ExtractionParameters ExtractImageParameters(OperationParameters parameters, string path)
    {
        var outputDir = parameters.GetOptional<string?>("outputDir") ?? Path.GetDirectoryName(path) ?? ".";
        var formatStr = parameters.GetOptional("format", "png");
        var skipDuplicates = parameters.GetOptional("skipDuplicates", false);

        var isJpeg = formatStr.ToLower() is "jpeg" or "jpg";
        var extension = isJpeg ? "jpg" : "png";

        return new ExtractionParameters(outputDir, isJpeg, extension, skipDuplicates);
    }

    /// <summary>
    ///     Extracts all images from the presentation.
    /// </summary>
    /// <param name="presentation">The presentation to extract images from.</param>
    /// <param name="p">The extraction parameters.</param>
    /// <returns>A tuple containing the extracted count and skipped count.</returns>
    private static (int count, int skippedCount) ExtractAllImages(Presentation presentation, ExtractionParameters p)
    {
        var count = 0;
        var skippedCount = 0;
        var exportedHashes = new HashSet<string>();

        var slideNum = 0;
        foreach (var slide in presentation.Slides)
        {
            slideNum++;
            foreach (var shape in slide.Shapes)
                if (shape is PictureFrame { PictureFormat.Picture.Image: not null } pic)
                {
                    var extracted = TryExtractImage(pic, slideNum, ref count, p, exportedHashes);
                    if (!extracted) skippedCount++;
                }
        }

        return (count, skippedCount);
    }

    /// <summary>
    ///     Tries to extract an image from a picture frame.
    /// </summary>
    /// <param name="pic">The picture frame.</param>
    /// <param name="slideNum">The slide number.</param>
    /// <param name="count">The current extraction count.</param>
    /// <param name="p">The extraction parameters.</param>
    /// <param name="exportedHashes">The set of exported image hashes.</param>
    /// <returns>True if the image was extracted, false if skipped.</returns>
    private static bool TryExtractImage(PictureFrame pic, int slideNum, ref int count,
        ExtractionParameters p, HashSet<string> exportedHashes)
    {
        var image = pic.PictureFormat.Picture.Image;

        if (p.SkipDuplicates)
        {
            var hash = PptImageHelper.ComputeImageHash(image.BinaryData);
            if (!exportedHashes.Add(hash))
                return false;
        }

        var fileName = Path.Combine(p.OutputDir, $"slide{slideNum}_img{++count}.{p.Extension}");
        image.SystemImage.Save(fileName, p.IsJpeg
            ? ImageFormat.Jpeg
            : ImageFormat.Png);
        return true;
    }

    /// <summary>
    ///     Builds the result message.
    /// </summary>
    /// <param name="count">The number of extracted images.</param>
    /// <param name="skippedCount">The number of skipped duplicates.</param>
    /// <param name="p">The extraction parameters.</param>
    /// <returns>The result message.</returns>
    private static string BuildResultMessage(int count, int skippedCount, ExtractionParameters p)
    {
        var result = $"Extracted {count} images. Output: {Path.GetFullPath(p.OutputDir)}";
        if (p.SkipDuplicates && skippedCount > 0)
            result += $" (skipped {skippedCount} duplicates)";
        return result;
    }

    /// <summary>
    ///     Record for holding extraction parameters.
    /// </summary>
    /// <param name="OutputDir">The output directory.</param>
    /// <param name="IsJpeg">Whether the output format is JPEG (false = PNG).</param>
    /// <param name="Extension">The file extension.</param>
    /// <param name="SkipDuplicates">Whether to skip duplicate images.</param>
    private sealed record ExtractionParameters(
        string OutputDir,
        bool IsJpeg,
        string Extension,
        bool SkipDuplicates);
}
