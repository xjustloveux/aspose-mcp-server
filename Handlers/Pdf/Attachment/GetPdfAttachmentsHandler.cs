using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

/// <summary>
///     Handler for retrieving attachments from PDF documents.
/// </summary>
public class GetPdfAttachmentsHandler : OperationHandlerBase<Document>
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Retrieves all attachments from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing attachment information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        var embeddedFiles = document.EmbeddedFiles;

        if (embeddedFiles == null || embeddedFiles.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                items = Array.Empty<object>(),
                message = "No attachments found"
            };
            return JsonSerializer.Serialize(emptyResult, JsonOptions);
        }

        var attachmentList = PdfAttachmentHelper.CollectAttachmentInfo(embeddedFiles);

        var result = new
        {
            count = attachmentList.Count,
            items = attachmentList
        };
        return JsonSerializer.Serialize(result, JsonOptions);
    }
}
