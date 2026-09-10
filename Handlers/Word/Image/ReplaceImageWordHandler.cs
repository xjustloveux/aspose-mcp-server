using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Word;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using IOFile = System.IO.File;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for replacing images in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractReplaceImageParameters(parameters);
        // A secondary input is still an input: the replacement image is read from disk, so it
        // has to come from inside the allowlist, and the reader must see the resolved path.
        var resolvedImagePath = ValidateParameters(p, context.ServerConfig?.AllowedBasePaths ?? []);

        var doc = context.Document;
        var allImages = WordImageHelper.GetAllImages(doc, p.SectionIndex);
        ValidateImageIndex(p.ImageIndex, allImages.Count);

        var shapeToReplace = allImages[p.ImageIndex];
        var originalProps = CaptureOriginalProperties(shapeToReplace, p.PreservePosition);

        ReplaceImage(shapeToReplace, p, originalProps, resolvedImagePath);

        MarkModified(context);

        return BuildResultMessage(p, originalProps.Width, originalProps.Height, shapeToReplace.Height);
    }

    /// <summary>
    ///     Validates required parameters and returns the allowlist-resolved image path.
    /// </summary>
    /// <param name="p">Extracted parameters.</param>
    /// <param name="allowedBasePaths">Roots the caller may read from.</param>
    /// <returns>The resolved path the image must be read from.</returns>
    private static string ValidateParameters(ReplaceImageParameters p,
        IReadOnlyList<string> allowedBasePaths)
    {
        if (string.IsNullOrEmpty(p.NewImagePath))
            throw new ArgumentException("newImagePath or imagePath is required for replace operation");

        SecurityHelper.ValidateFilePath(p.NewImagePath, "newImagePath", true);
        var resolvedImagePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.NewImagePath,
            allowedBasePaths, "newImagePath");

        if (!IOFile.Exists(resolvedImagePath))
            throw new FileNotFoundException("The specified file was not found.");

        return resolvedImagePath;
    }

    /// <summary>
    ///     Validates image index is within range.
    /// </summary>
    private static void ValidateImageIndex(int imageIndex, int imageCount)
    {
        if (imageIndex < 0 || imageIndex >= imageCount)
            throw new ArgumentException(
                $"Image index {imageIndex} is out of range (document has {imageCount} images)");
    }

    /// <summary>
    ///     Captures original properties from the shape.
    /// </summary>
    /// <returns>The original shape properties.</returns>
    private static OriginalShapeProperties CaptureOriginalProperties(WordShape shape, bool capturePosition)
    {
        return new OriginalShapeProperties(
            shape.Width,
            shape.Height,
            shape.WrapType,
            capturePosition ? shape.HorizontalAlignment : null,
            capturePosition ? shape.VerticalAlignment : null,
            capturePosition ? shape.RelativeHorizontalPosition : null,
            capturePosition ? shape.RelativeVerticalPosition : null,
            capturePosition ? shape.Left : null,
            capturePosition ? shape.Top : null);
    }

    /// <summary>
    ///     Replaces the image and applies size/position settings.
    /// </summary>
    private static void ReplaceImage(WordShape shape, ReplaceImageParameters p,
        OriginalShapeProperties originalProps, string resolvedImagePath)
    {
        try
        {
            shape.ImageData.SetImage(resolvedImagePath);
            ApplySizeSettings(shape, p, originalProps);
            ApplyPositionSettings(shape, p, originalProps);
        }
        catch (Exception ex)
        {
            throw WordErrorTranslator.Translate(ex);
        }
    }

    /// <summary>
    ///     Applies size settings based on parameters.
    /// </summary>
    private static void ApplySizeSettings(WordShape shape, ReplaceImageParameters p,
        OriginalShapeProperties originalProps)
    {
        if (!p.PreserveSize) return;

        if (p.SmartFit)
            ApplySmartFitSize(shape, originalProps.Width, originalProps.Height);
        else
            ApplyOriginalSize(shape, originalProps.Width, originalProps.Height);
    }

    /// <summary>
    ///     Applies smart fit size calculation.
    /// </summary>
    private static void ApplySmartFitSize(WordShape shape, double originalWidth, double originalHeight)
    {
        var newImageSize = shape.ImageData.ImageSize;
        if (newImageSize.WidthPixels > 0)
        {
            var newAspectRatio = (double)newImageSize.HeightPixels / newImageSize.WidthPixels;
            shape.Width = originalWidth;
            shape.Height = originalWidth * newAspectRatio;
        }
        else
        {
            ApplyOriginalSize(shape, originalWidth, originalHeight);
        }
    }

    /// <summary>
    ///     Applies original size to the shape.
    /// </summary>
    private static void ApplyOriginalSize(WordShape shape, double width, double height)
    {
        shape.Width = width;
        shape.Height = height;
    }

    /// <summary>
    ///     Applies position settings based on parameters.
    /// </summary>
    private static void ApplyPositionSettings(WordShape shape, ReplaceImageParameters p,
        OriginalShapeProperties originalProps)
    {
        if (!p.PreservePosition) return;

        shape.WrapType = originalProps.WrapType;

        if (originalProps.HorizontalAlignment.HasValue)
            shape.HorizontalAlignment = originalProps.HorizontalAlignment.Value;
        if (originalProps.VerticalAlignment.HasValue)
            shape.VerticalAlignment = originalProps.VerticalAlignment.Value;
        if (originalProps.RelativeHorizontalPosition.HasValue)
            shape.RelativeHorizontalPosition = originalProps.RelativeHorizontalPosition.Value;
        if (originalProps.RelativeVerticalPosition.HasValue)
            shape.RelativeVerticalPosition = originalProps.RelativeVerticalPosition.Value;
        if (originalProps.Left.HasValue)
            shape.Left = originalProps.Left.Value;
        if (originalProps.Top.HasValue)
            shape.Top = originalProps.Top.Value;
    }

    /// <summary>
    ///     Builds the result message.
    /// </summary>
    /// <returns>The success result.</returns>
    private static SuccessResult BuildResultMessage(ReplaceImageParameters p, double originalWidth,
        double originalHeight, double newHeight)
    {
        var result = $"Image #{p.ImageIndex} replaced successfully\n";
        result += $"New image: {Path.GetFileName(p.NewImagePath)}\n";

        if (p.PreserveSize)
            result += p.SmartFit
                ? $"Smart fit: width preserved ({originalWidth:F1} pt), height calculated proportionally ({newHeight:F1} pt)\n"
                : $"Preserved size: {originalWidth:F1} pt x {originalHeight:F1} pt\n";

        if (p.PreservePosition)
            result += "Preserved position and wrapping";

        return new SuccessResult { Message = result.TrimEnd() };
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

    private sealed record ReplaceImageParameters(
        int ImageIndex,
        string? NewImagePath,
        bool PreserveSize,
        bool SmartFit,
        bool PreservePosition,
        int SectionIndex);

    private sealed record OriginalShapeProperties(
        double Width,
        double Height,
        WrapType WrapType,
        HorizontalAlignment? HorizontalAlignment,
        VerticalAlignment? VerticalAlignment,
        RelativeHorizontalPosition? RelativeHorizontalPosition,
        RelativeVerticalPosition? RelativeVerticalPosition,
        double? Left,
        double? Top);
}
