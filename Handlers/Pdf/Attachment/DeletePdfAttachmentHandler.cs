using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Pdf;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

/// <summary>
///     Handler for deleting attachments from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        var document = context.Document;
        var embeddedFiles = document.EmbeddedFiles;

        var (found, actualName, attachmentNames) =
            PdfAttachmentHelper.FindAttachment(embeddedFiles, deleteParams.AttachmentName);

        if (!found)
        {
            var availableNames = string.Join(", ", attachmentNames);
            throw new ArgumentException(
                $"Attachment '{deleteParams.AttachmentName}' not found. Available attachments: {(string.IsNullOrEmpty(availableNames) ? "(none)" : availableNames)}");
        }

        embeddedFiles.Delete(actualName);

        MarkModified(context);

        return new SuccessResult { Message = $"Deleted attachment '{deleteParams.AttachmentName}'." };
    }

    /// <summary>
    ///     Extracts delete parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<string>("attachmentName")
        );
    }

    /// <summary>
    ///     Record to hold delete attachment parameters.
    /// </summary>
    private sealed record DeleteParameters(string AttachmentName);
}
