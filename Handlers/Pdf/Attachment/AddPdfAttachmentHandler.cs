using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

/// <summary>
///     Handler for adding attachments to PDF documents.
/// </summary>
public class AddPdfAttachmentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a file attachment to the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: attachmentPath, attachmentName
    ///     Optional: description
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var attachmentPath = parameters.GetRequired<string>("attachmentPath");
        var attachmentName = parameters.GetRequired<string>("attachmentName");
        var description = parameters.GetOptional<string?>("description");

        SecurityHelper.ValidateFilePath(attachmentPath, "attachmentPath", true);
        SecurityHelper.ValidateStringLength(attachmentName, "attachmentName", 255);
        if (description != null)
            SecurityHelper.ValidateStringLength(description, "description", 1000);

        if (!File.Exists(attachmentPath))
            throw new FileNotFoundException($"Attachment file not found: {attachmentPath}");

        var document = context.Document;
        var existingNames = PdfAttachmentHelper.CollectAttachmentNames(document.EmbeddedFiles);
        if (existingNames.Any(n => string.Equals(n, attachmentName, StringComparison.OrdinalIgnoreCase) ||
                                   string.Equals(Path.GetFileName(n), attachmentName,
                                       StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Attachment with name '{attachmentName}' already exists");

        var fileSpecification = new FileSpecification(attachmentPath, description ?? "")
        {
            Name = attachmentName
        };

        document.EmbeddedFiles.Add(fileSpecification);

        MarkModified(context);

        return Success($"Added attachment '{attachmentName}'.");
    }
}
