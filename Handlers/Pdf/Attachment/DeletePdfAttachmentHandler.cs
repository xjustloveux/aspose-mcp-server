using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

/// <summary>
///     Handler for deleting attachments from PDF documents.
/// </summary>
public class DeletePdfAttachmentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes an attachment from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: attachmentName
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var attachmentName = parameters.GetRequired<string>("attachmentName");

        var document = context.Document;
        var embeddedFiles = document.EmbeddedFiles;

        var (found, actualName, attachmentNames) = PdfAttachmentHelper.FindAttachment(embeddedFiles, attachmentName);

        if (!found)
        {
            var availableNames = string.Join(", ", attachmentNames);
            throw new ArgumentException(
                $"Attachment '{attachmentName}' not found. Available attachments: {(string.IsNullOrEmpty(availableNames) ? "(none)" : availableNames)}");
        }

        embeddedFiles.Delete(actualName);

        MarkModified(context);

        return Success($"Deleted attachment '{attachmentName}'.");
    }
}
