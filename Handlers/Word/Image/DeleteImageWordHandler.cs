using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for deleting images from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteImageParameters(parameters);

        var doc = context.Document;

        var allImages = WordImageHelper.GetAllImages(doc, p.SectionIndex);

        if (p.ImageIndex < 0 || p.ImageIndex >= allImages.Count)
            throw new ArgumentException(
                $"Image index {p.ImageIndex} is out of range (document has {allImages.Count} images)");

        var shapeToDelete = allImages[p.ImageIndex];

        var imageInfo = $"Image #{p.ImageIndex}";
        if (shapeToDelete.HasImage)
            try
            {
                imageInfo += $" (Width: {shapeToDelete.Width:F1} pt, Height: {shapeToDelete.Height:F1} pt)";
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Failed to get image size information: {ex.Message}");
            }

        shapeToDelete.Remove();

        MarkModified(context);

        var remainingCount = WordImageHelper.GetAllImages(doc, p.SectionIndex).Count;

        return new SuccessResult { Message = $"{imageInfo} deleted successfully\nRemaining images: {remainingCount}" };
    }

    /// <summary>
    ///     Extracts delete image parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete image parameters.</returns>
    private static DeleteImageParameters ExtractDeleteImageParameters(OperationParameters parameters)
    {
        return new DeleteImageParameters(
            parameters.GetOptional("imageIndex", 0),
            parameters.GetOptional("sectionIndex", 0)
        );
    }

    /// <summary>
    ///     Record to hold delete image parameters.
    /// </summary>
    private sealed record DeleteImageParameters(int ImageIndex, int SectionIndex);
}
