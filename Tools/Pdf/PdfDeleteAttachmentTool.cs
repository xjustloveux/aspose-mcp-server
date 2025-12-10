using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfDeleteAttachmentTool : IAsposeTool
{
    public string Description => "Delete an attachment from PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            attachmentName = new
            {
                type = "string",
                description = "Attachment name"
            }
        },
        required = new[] { "path", "attachmentName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var attachmentName = arguments?["attachmentName"]?.GetValue<string>() ?? throw new ArgumentException("attachmentName is required");

        using var document = new Document(path);
        var attachments = document.EmbeddedFiles;
        
        if (attachments == null || attachments.Count == 0)
        {
            throw new ArgumentException("No attachments found in the document");
        }

        FileSpecification? attachmentToDelete = null;
        for (int i = 0; i < attachments.Count; i++)
        {
            var attachment = attachments[i];
            if (attachment.Name == attachmentName)
            {
                attachmentToDelete = attachment;
                break;
            }
        }

        if (attachmentToDelete == null)
        {
            throw new ArgumentException($"Attachment '{attachmentName}' not found");
        }

        // Remove attachment by name
        attachments.Delete(attachmentName);
        document.Save(path);

        return await Task.FromResult($"Attachment '{attachmentName}' deleted: {path}");
    }
}

