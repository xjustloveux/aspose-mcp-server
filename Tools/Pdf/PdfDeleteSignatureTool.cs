using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;

namespace AsposeMcpServer.Tools;

public class PdfDeleteSignatureTool : IAsposeTool
{
    public string Description => "Delete a digital signature from PDF document";

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
            signatureName = new
            {
                type = "string",
                description = "Signature field name (PartialName or FullName)"
            }
        },
        required = new[] { "path", "signatureName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var signatureName = arguments?["signatureName"]?.GetValue<string>() ?? throw new ArgumentException("signatureName is required");

        using var document = new Document(path);
        if (document.Form == null)
        {
            throw new ArgumentException("No form fields found in the document");
        }

        SignatureField? signatureToDelete = null;
        foreach (Field field in document.Form)
        {
            if (field is SignatureField signature && (signature.PartialName == signatureName || signature.FullName == signatureName))
            {
                signatureToDelete = signature;
                break;
            }
        }

        if (signatureToDelete == null)
        {
            throw new ArgumentException($"Signature '{signatureName}' not found");
        }

        document.Form.Delete(signatureToDelete);
        document.Save(path);

        return await Task.FromResult($"Signature '{signatureName}' deleted: {path}");
    }
}

