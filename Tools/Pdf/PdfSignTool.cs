using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using Aspose.Pdf.Facades;

namespace AsposeMcpServer.Tools;

public class PdfSignTool : IAsposeTool
{
    public string Description => "Digitally sign a PDF document";

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
            certificatePath = new
            {
                type = "string",
                description = "Path to certificate file (.pfx)"
            },
            certificatePassword = new
            {
                type = "string",
                description = "Certificate password"
            },
            reason = new
            {
                type = "string",
                description = "Reason for signing (optional)"
            },
            location = new
            {
                type = "string",
                description = "Location of signing (optional)"
            }
        },
        required = new[] { "path", "certificatePath", "certificatePassword" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var certificatePath = arguments?["certificatePath"]?.GetValue<string>() ?? throw new ArgumentException("certificatePath is required");
        var certificatePassword = arguments?["certificatePassword"]?.GetValue<string>() ?? throw new ArgumentException("certificatePassword is required");
        var reason = arguments?["reason"]?.GetValue<string>() ?? "Document approval";
        var location = arguments?["location"]?.GetValue<string>() ?? "";

        using var document = new Document(path);
        
        using var pdfSign = new PdfFileSignature(document);
        var pkcs = new Aspose.Pdf.Forms.PKCS7(certificatePath, certificatePassword);
        pkcs.Reason = reason;
        pkcs.Location = location;

        var rect = new System.Drawing.Rectangle(100, 100, 200, 100);
        pdfSign.Sign(1, true, rect, pkcs);
        pdfSign.Save(path);

        return await Task.FromResult($"PDF digitally signed: {path}");
    }
}

