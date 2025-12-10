using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfGetDocumentPropertiesTool : IAsposeTool
{
    public string Description => "Get document properties (metadata) from PDF file";

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
        var metadata = document.Metadata;
        var sb = new StringBuilder();

        sb.AppendLine("Document Properties:");
        sb.AppendLine($"  Title: {metadata["Title"] ?? "(none)"}");
        sb.AppendLine($"  Author: {metadata["Author"] ?? "(none)"}");
        sb.AppendLine($"  Subject: {metadata["Subject"] ?? "(none)"}");
        sb.AppendLine($"  Keywords: {metadata["Keywords"] ?? "(none)"}");
        sb.AppendLine($"  Creator: {metadata["Creator"] ?? "(none)"}");
        sb.AppendLine($"  Producer: {metadata["Producer"] ?? "(none)"}");
        sb.AppendLine($"  Creation Date: {metadata["CreationDate"] ?? "(none)"}");
        sb.AppendLine($"  Modification Date: {metadata["ModDate"] ?? "(none)"}");
        sb.AppendLine();
        sb.AppendLine($"Total Pages: {document.Pages.Count}");
        sb.AppendLine($"Is Encrypted: {document.IsEncrypted}");
        sb.AppendLine($"Is Linearized: {document.IsLinearized}");

        return await Task.FromResult(sb.ToString());
    }
}

