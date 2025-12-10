using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfGetAttachmentsTool : IAsposeTool
{
    public string Description => "Get all attachments from PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var document = new Document(path);
        var sb = new StringBuilder();

        var attachments = document.EmbeddedFiles;
        if (attachments == null || attachments.Count == 0)
        {
            return await Task.FromResult("No attachments found in the document.");
        }

        sb.AppendLine($"Attachments ({attachments.Count}):");
        sb.AppendLine();

        for (int i = 0; i < attachments.Count; i++)
        {
            var attachment = attachments[i];
            sb.AppendLine($"[{i}] {attachment.Name ?? "(unnamed)"}");
            sb.AppendLine($"  Description: {attachment.Description ?? "(none)"}");
            sb.AppendLine($"  MIME Type: {attachment.MIMEType ?? "(unknown)"}");
            if (attachment.Params != null)
            {
                sb.AppendLine($"  Size: {attachment.Params.Size} bytes");
                sb.AppendLine($"  Creation Date: {attachment.Params.CreationDate}");
                sb.AppendLine($"  Modification Date: {attachment.Params.ModDate}");
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

