using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for extracting embedded images from PowerPoint presentations.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var path = context.SourcePath;
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for extract operation");

        SecurityHelper.ValidateFilePath(path, "path", true);

        var outputDir = parameters.GetOptional<string?>("outputDir");
        var formatStr = parameters.GetOptional("format", "png");
        var skipDuplicates = parameters.GetOptional("skipDuplicates", false);

        var actualOutputDir = outputDir ?? Path.GetDirectoryName(path) ?? ".";

#pragma warning disable CA1416
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416

        Directory.CreateDirectory(actualOutputDir);

        var presentation = context.Document;
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
                    var image = pic.PictureFormat.Picture.Image;

                    if (skipDuplicates)
                    {
                        var hash = PptImageHelper.ComputeImageHash(image.BinaryData);
                        if (!exportedHashes.Add(hash))
                        {
                            skippedCount++;
                            continue;
                        }
                    }

                    var fileName = Path.Combine(actualOutputDir, $"slide{slideNum}_img{++count}.{extension}");
                    var systemImage = image.SystemImage;
#pragma warning disable CA1416
                    systemImage.Save(fileName, format);
#pragma warning restore CA1416
                }
        }

        var result = $"Extracted {count} images. Output: {Path.GetFullPath(actualOutputDir)}";
        if (skipDuplicates && skippedCount > 0)
            result += $" (skipped {skippedCount} duplicates)";

        return Success(result);
    }
}
