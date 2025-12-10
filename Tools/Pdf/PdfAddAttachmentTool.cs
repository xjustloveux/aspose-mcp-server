using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using System.IO;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfAddAttachmentTool : IAsposeTool
{
    public string Description => "Add an attachment (file) to PDF document";

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
            attachmentPath = new
            {
                type = "string",
                description = "Path to file to attach"
            },
            attachmentName = new
            {
                type = "string",
                description = "Attachment name (optional, defaults to filename)"
            },
            description = new
            {
                type = "string",
                description = "Attachment description (optional)"
            }
        },
        required = new[] { "path", "attachmentPath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var attachmentPath = arguments?["attachmentPath"]?.GetValue<string>() ?? throw new ArgumentException("attachmentPath is required");
        var attachmentName = arguments?["attachmentName"]?.GetValue<string>();
        var description = arguments?["description"]?.GetValue<string>();

        if (!File.Exists(attachmentPath))
        {
            throw new FileNotFoundException($"Attachment file not found: {attachmentPath}");
        }

        using var document = new Document(path);
        var fileSpecification = new FileSpecification(attachmentPath);
        
        if (!string.IsNullOrEmpty(attachmentName))
        {
            fileSpecification.Name = SecurityHelper.SanitizeFileName(attachmentName);
        }
        
        if (!string.IsNullOrEmpty(description))
        {
            fileSpecification.Description = description;
        }

        document.EmbeddedFiles.Add(fileSpecification);
        document.Save(path);

        return await Task.FromResult($"Attachment '{fileSpecification.Name}' added: {path}");
    }
}

