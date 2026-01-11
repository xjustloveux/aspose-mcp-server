using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for deleting images from Word documents.
/// </summary>
public class DeleteImageWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes an image from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imageIndex
    ///     Optional: sectionIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var imageIndex = parameters.GetOptional("imageIndex", 0);
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);

        var doc = context.Document;

        var allImages = WordImageHelper.GetAllImages(doc, sectionIndex);

        if (imageIndex < 0 || imageIndex >= allImages.Count)
            throw new ArgumentException(
                $"Image index {imageIndex} is out of range (document has {allImages.Count} images)");

        var shapeToDelete = allImages[imageIndex];

        var imageInfo = $"Image #{imageIndex}";
        if (shapeToDelete.HasImage)
            try
            {
                imageInfo += $" (Width: {shapeToDelete.Width:F1} pt, Height: {shapeToDelete.Height:F1} pt)";
            }
            catch (Exception ex)
            {
                // Size information may not be available, but this is not critical
                Console.Error.WriteLine($"[WARN] Failed to get image size information: {ex.Message}");
            }

        shapeToDelete.Remove();

        MarkModified(context);

        var remainingCount = WordImageHelper.GetAllImages(doc, sectionIndex).Count;

        return $"{imageInfo} deleted successfully\nRemaining images: {remainingCount}";
    }
}
