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
        var imageIndex = parameters.GetOptional("imageIndex", 0);
        var newImagePath = parameters.GetOptional<string?>("newImagePath") ??
                           parameters.GetOptional<string?>("imagePath");
        var preserveSize = parameters.GetOptional("preserveSize", true);
        var smartFit = parameters.GetOptional("smartFit", false);
        var preservePosition = parameters.GetOptional("preservePosition", true);
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);

        if (string.IsNullOrEmpty(newImagePath))
            throw new ArgumentException("newImagePath or imagePath is required for replace operation");

        SecurityHelper.ValidateFilePath(newImagePath, "newImagePath", true);

        if (!IOFile.Exists(newImagePath))
            throw new FileNotFoundException($"Image file not found: {newImagePath}");

        var doc = context.Document;

        var allImages = WordImageHelper.GetAllImages(doc, sectionIndex);

        if (imageIndex < 0 || imageIndex >= allImages.Count)
            throw new ArgumentException(
                $"Image index {imageIndex} is out of range (document has {allImages.Count} images)");

        var shapeToReplace = allImages[imageIndex];

        var originalWidth = shapeToReplace.Width;
        var originalHeight = shapeToReplace.Height;
        var originalWrapType = shapeToReplace.WrapType;
        HorizontalAlignment? originalHorizontalAlignment = null;
        VerticalAlignment? originalVerticalAlignment = null;
        RelativeHorizontalPosition? originalRelativeHorizontalPosition = null;
        RelativeVerticalPosition? originalRelativeVerticalPosition = null;
        double? originalLeft = null;
        double? originalTop = null;

        if (preservePosition)
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
            shapeToReplace.ImageData.SetImage(newImagePath);

            if (preserveSize)
            {
                if (smartFit)
                {
                    // Calculate proportional height based on new image's aspect ratio
                    var newImageSize = shapeToReplace.ImageData.ImageSize;
                    if (newImageSize.WidthPixels > 0)
                    {
                        var newAspectRatio = (double)newImageSize.HeightPixels / newImageSize.WidthPixels;
                        shapeToReplace.Width = originalWidth;
                        shapeToReplace.Height = originalWidth * newAspectRatio;
                    }
                    else
                    {
                        // Fallback to original size if aspect ratio can't be calculated
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

            if (preservePosition)
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

        var result = $"Image #{imageIndex} replaced successfully\n";
        result += $"New image: {Path.GetFileName(newImagePath)}\n";
        if (preserveSize)
        {
            if (smartFit)
                result +=
                    $"Smart fit: width preserved ({originalWidth:F1} pt), height calculated proportionally ({shapeToReplace.Height:F1} pt)\n";
            else
                result += $"Preserved size: {originalWidth:F1} pt x {originalHeight:F1} pt\n";
        }

        if (preservePosition) result += "Preserved position and wrapping";

        return result.TrimEnd();
    }
}
