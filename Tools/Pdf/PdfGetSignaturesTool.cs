using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Forms;

namespace AsposeMcpServer.Tools;

public class PdfGetSignaturesTool : IAsposeTool
{
    public string Description => "Get all digital signatures from PDF document";

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

        if (document.Form == null)
        {
            return await Task.FromResult("No signatures found in the document.");
        }

        var signatures = new List<SignatureField>();
        foreach (Field field in document.Form)
        {
            if (field is SignatureField signature)
            {
                signatures.Add(signature);
            }
        }

        if (signatures.Count == 0)
        {
            return await Task.FromResult("No signatures found in the document.");
        }

        sb.AppendLine($"Signatures ({signatures.Count}):");
        sb.AppendLine();

        for (int i = 0; i < signatures.Count; i++)
        {
            var signature = signatures[i];
            sb.AppendLine($"[{i}] Signature:");
            sb.AppendLine($"  Name: {signature.PartialName ?? "(unnamed)"}");
            sb.AppendLine($"  Full Name: {signature.FullName ?? "(none)"}");
            sb.AppendLine($"  Position: ({signature.Rect.LLX}, {signature.Rect.LLY})");
            sb.AppendLine($"  Size: ({signature.Rect.Width}, {signature.Rect.Height})");
            if (signature.Signature != null)
            {
                sb.AppendLine($"  Reason: {signature.Signature.Reason ?? "(none)"}");
                sb.AppendLine($"  Location: {signature.Signature.Location ?? "(none)"}");
                sb.AppendLine($"  Contact Info: {signature.Signature.ContactInfo ?? "(none)"}");
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

