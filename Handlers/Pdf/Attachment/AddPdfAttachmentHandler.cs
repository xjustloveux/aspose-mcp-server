using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Pdf;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

/// <summary>
///     Handler for adding attachments to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var addParams = ExtractAddParameters(parameters);

        SecurityHelper.ValidateFilePath(addParams.AttachmentPath, "attachmentPath", true);
        SecurityHelper.ValidateStringLength(addParams.AttachmentName, "attachmentName", 255);
        if (addParams.Description != null)
            SecurityHelper.ValidateStringLength(addParams.Description, "description", 1000);

        if (!File.Exists(addParams.AttachmentPath))
            throw new FileNotFoundException($"Attachment file not found: {addParams.AttachmentPath}");

        var document = context.Document;
        var existingNames = PdfAttachmentHelper.CollectAttachmentNames(document.EmbeddedFiles);
        if (existingNames.Any(n => string.Equals(n, addParams.AttachmentName, StringComparison.OrdinalIgnoreCase) ||
                                   string.Equals(Path.GetFileName(n), addParams.AttachmentName,
                                       StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Attachment with name '{addParams.AttachmentName}' already exists");

        var fileSpecification = new FileSpecification(addParams.AttachmentPath, addParams.Description ?? "")
        {
            Name = addParams.AttachmentName
        };

        document.EmbeddedFiles.Add(fileSpecification);

        MarkModified(context);

        return new SuccessResult { Message = $"Added attachment '{addParams.AttachmentName}'." };
    }

    /// <summary>
    ///     Extracts add parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<string>("attachmentPath"),
            parameters.GetRequired<string>("attachmentName"),
            parameters.GetOptional<string?>("description")
        );
    }

    /// <summary>
    ///     Record to hold add attachment parameters.
    /// </summary>
    private sealed record AddParameters(string AttachmentPath, string AttachmentName, string? Description);
}
