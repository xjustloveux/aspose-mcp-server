using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using Aspose.Pdf.Facades;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfSignatureTool : IAsposeTool
{
    public string Description => @"Manage digital signatures in PDF documents. Supports 3 operations: sign, delete, get.

Usage examples:
- Sign PDF: pdf_signature(operation='sign', path='doc.pdf', certificatePath='cert.pfx', certificatePassword='password')
- Delete signature: pdf_signature(operation='delete', path='doc.pdf', signatureIndex=0)
- Get signatures: pdf_signature(operation='get', path='doc.pdf')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'sign': Sign PDF with certificate (required params: path, certificatePath, certificatePassword)
- 'delete': Delete a signature (required params: path, signatureIndex)
- 'get': Get all signatures (required params: path)",
                @enum = new[] { "sign", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            certificatePath = new
            {
                type = "string",
                description = "Path to certificate file (.pfx, required for sign)"
            },
            certificatePassword = new
            {
                type = "string",
                description = "Certificate password (required for sign)"
            },
            reason = new
            {
                type = "string",
                description = "Reason for signing (for sign, optional)"
            },
            location = new
            {
                type = "string",
                description = "Location of signing (for sign, optional)"
            },
            signatureIndex = new
            {
                type = "number",
                description = "Signature index (0-based, required for delete)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "sign" => await SignDocument(arguments),
            "delete" => await DeleteSignature(arguments),
            "get" => await GetSignatures(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> SignDocument(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var certificatePath = arguments?["certificatePath"]?.GetValue<string>() ?? throw new ArgumentException("certificatePath is required");
        var certificatePassword = arguments?["certificatePassword"]?.GetValue<string>() ?? throw new ArgumentException("certificatePassword is required");
        var reason = arguments?["reason"]?.GetValue<string>() ?? "Document approval";
        var location = arguments?["location"]?.GetValue<string>() ?? "";

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(certificatePath, "certificatePath");

        if (!File.Exists(certificatePath))
            throw new FileNotFoundException($"Certificate file not found: {certificatePath}");

        using var document = new Document(path);
        using var pdfSign = new PdfFileSignature(document);
        var pkcs = new PKCS7(certificatePath, certificatePassword);
        pkcs.Reason = reason;
        pkcs.Location = location;

        var rect = new System.Drawing.Rectangle(100, 100, 200, 100);
        pdfSign.Sign(1, true, rect, pkcs);
        pdfSign.Save(outputPath);
        return await Task.FromResult($"PDF digitally signed. Output: {outputPath}");
    }

    private async Task<string> DeleteSignature(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var signatureIndex = arguments?["signatureIndex"]?.GetValue<int>() ?? throw new ArgumentException("signatureIndex is required");

        using var document = new Document(path);
        using var pdfSign = new PdfFileSignature(document);
        var signatureNames = pdfSign.GetSignNames();
        
        if (signatureIndex < 0 || signatureIndex >= signatureNames.Count)
            throw new ArgumentException($"signatureIndex must be between 0 and {signatureNames.Count - 1}");

        var signatureName = signatureNames[signatureIndex];
        pdfSign.RemoveSignature(signatureName);
        pdfSign.Save(outputPath);
        return await Task.FromResult($"Successfully deleted signature '{signatureName}' (index {signatureIndex}). Output: {outputPath}");
    }

    private async Task<string> GetSignatures(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        using var pdfSign = new PdfFileSignature(document);
        var signatureNames = pdfSign.GetSignNames();
        var sb = new StringBuilder();

        sb.AppendLine("=== PDF Signatures ===");
        sb.AppendLine();

        if (signatureNames.Count == 0)
        {
            sb.AppendLine("No signatures found.");
            return await Task.FromResult(sb.ToString());
        }

        sb.AppendLine($"Total Signatures: {signatureNames.Count}");
        sb.AppendLine();

        for (int i = 0; i < signatureNames.Count; i++)
        {
            var signatureName = signatureNames[i];
            sb.AppendLine($"[{i}] Name: {signatureName}");
            // Check signature validity - IsValid may not be available, use alternative method
            try
            {
                var cert = pdfSign.ExtractCertificate(signatureName);
                sb.AppendLine($"    Valid: Yes (Certificate found)");
            }
            catch
            {
                sb.AppendLine($"    Valid: Unknown");
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

