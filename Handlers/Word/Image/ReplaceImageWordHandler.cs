using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using IOFile = System.IO.File;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for replacing images in Word documents.
/// </summary>
public class ReplaceImageWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "replace";

    /// <summary>
    ///     Replaces an existing image with a new one.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imageIndex, newImagePath (or imagePath)
    ///     Optional: preserveSize, smartFit, preservePosition, sectionIndex
    /// </param>
    /// <returns>Success message with replacement details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractReplaceImageParameters(parameters);

        if (string.IsNullOrEmpty(p.NewImagePath))
            throw new ArgumentException("newImagePath or imagePath is required for replace operation");

        SecurityHelper.ValidateFilePath(p.NewImagePath, "newImagePath", true);

        if (!IOFile.Exists(p.NewImagePath))
            throw new FileNotFoundException($"Image file not found: {p.NewImagePath}");

        var doc = context.Document;

        var allImages = WordImageHelper.GetAllImages(doc, p.SectionIndex);

        if (p.ImageIndex < 0 || p.ImageIndex >= allImages.Count)
            throw new ArgumentException(
                $"Image index {p.ImageIndex} is out of range (document has {allImages.Count} images)");

        var shapeToReplace = allImages[p.ImageIndex];

        var originalWidth = shapeToReplace.Width;
        var originalHeight = shapeToReplace.Height;
        var originalWrapType = shapeToReplace.WrapType;
        HorizontalAlignment? originalHorizontalAlignment = null;
        VerticalAlignment? originalVerticalAlignment = null;
        RelativeHorizontalPosition? originalRelativeHorizontalPosition = null;
        RelativeVerticalPosition? originalRelativeVerticalPosition = null;
        double? originalLeft = null;
        double? originalTop = null;

        if (p.PreservePosition)
        {
            originalHorizontalAlignment = shapeToReplace.HorizontalAlignment;
            originalVerticalAlignment = shapeToReplace.VerticalAlignment;
            originalRelativeHorizontalPosition = shapeToReplace.RelativeHorizontalPosition;
            originalRelativeVerticalPosition = shapeToReplace.RelativeVerticalPosition;
            originalLeft = shapeToReplace.Left;
            originalTop = shapeToReplace.Top;
        }

        try
        {
            shapeToReplace.ImageData.SetImage(p.NewImagePath);

            if (p.PreserveSize)
            {
                if (p.SmartFit)
                {
                    var newImageSize = shapeToReplace.ImageData.ImageSize;
                    if (newImageSize.WidthPixels > 0)
                    {
                        var newAspectRatio = (double)newImageSize.HeightPixels / newImageSize.WidthPixels;
                        shapeToReplace.Width = originalWidth;
                        shapeToReplace.Height = originalWidth * newAspectRatio;
                    }
                    else
                    {
                        shapeToReplace.Width = originalWidth;
                        shapeToReplace.Height = originalHeight;
                    }
                }
                else
                {
                    shapeToReplace.Width = originalWidth;
                    shapeToReplace.Height = originalHeight;
                }
            }

            if (p.PreservePosition)
            {
                shapeToReplace.WrapType = originalWrapType;
                if (originalHorizontalAlignment.HasValue)
                    shapeToReplace.HorizontalAlignment = originalHorizontalAlignment.Value;
                if (originalVerticalAlignment.HasValue)
                    shapeToReplace.VerticalAlignment = originalVerticalAlignment.Value;
                if (originalRelativeHorizontalPosition.HasValue)
                    shapeToReplace.RelativeHorizontalPosition = originalRelativeHorizontalPosition.Value;
                if (originalRelativeVerticalPosition.HasValue)
                    shapeToReplace.RelativeVerticalPosition = originalRelativeVerticalPosition.Value;
                if (originalLeft.HasValue)
                    shapeToReplace.Left = originalLeft.Value;
                if (originalTop.HasValue)
                    shapeToReplace.Top = originalTop.Value;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error occurred while replacing image: {ex.Message}", ex);
        }

        MarkModified(context);

        var result = $"Image #{p.ImageIndex} replaced successfully\n";
        result += $"New image: {Path.GetFileName(p.NewImagePath)}\n";
        if (p.PreserveSize)
        {
            if (p.SmartFit)
                result +=
                    $"Smart fit: width preserved ({originalWidth:F1} pt), height calculated proportionally ({shapeToReplace.Height:F1} pt)\n";
            else
                result += $"Preserved size: {originalWidth:F1} pt x {originalHeight:F1} pt\n";
        }

        if (p.PreservePosition) result += "Preserved position and wrapping";

        return result.TrimEnd();
    }

    private static ReplaceImageParameters ExtractReplaceImageParameters(OperationParameters parameters)
    {
        return new ReplaceImageParameters(
            parameters.GetOptional("imageIndex", 0),
            parameters.GetOptional<string?>("newImagePath") ?? parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional("preserveSize", true),
            parameters.GetOptional("smartFit", false),
            parameters.GetOptional("preservePosition", true),
            parameters.GetOptional("sectionIndex", 0));
    }

    private record ReplaceImageParameters(
        int ImageIndex,
        string? NewImagePath,
        bool PreserveSize,
        bool SmartFit,
        bool PreservePosition,
        int SectionIndex);
}
