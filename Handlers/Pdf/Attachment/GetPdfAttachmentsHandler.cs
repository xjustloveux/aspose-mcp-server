using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Pdf;
using AsposeMcpServer.Results.Pdf.Attachment;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

/// <summary>
///     Handler for retrieving attachments from PDF documents.
/// </summary>
[ResultType(typeof(GetAttachmentsResult))]
public class GetPdfAttachmentsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Retrieves all attachments from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing attachment information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        var embeddedFiles = document.EmbeddedFiles;

        if (embeddedFiles == null || embeddedFiles.Count == 0)
        {
            var emptyResult = new GetAttachmentsResult
            {
                Count = 0,
                Items = Array.Empty<AttachmentInfo>(),
                Message = "No attachments found"
            };
            return emptyResult;
        }

        var attachmentList = PdfAttachmentHelper.CollectAttachmentInfo(embeddedFiles);

        var result = new GetAttachmentsResult
        {
            Count = attachmentList.Count,
            Items = attachmentList
        };
        return result;
    }
}
