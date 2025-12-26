using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;
using Rectangle = System.Drawing.Rectangle;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing digital signatures in PDF documents (sign, delete, get)
/// </summary>
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        // Only get outputPath for operations that modify the document
        string? outputPath = null;
        if (operation.ToLower() != "get")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "sign" => await SignDocument(path, outputPath!, arguments),
            "delete" => await DeleteSignature(path, outputPath!, arguments),
            "get" => await GetSignatures(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Signs the PDF document
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing certificatePath, certificatePassword, optional reason, location</param>
    /// <returns>Success message</returns>
    private Task<string> SignDocument(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var certificatePath = ArgumentHelper.GetString(arguments, "certificatePath");
            var certificatePassword = ArgumentHelper.GetString(arguments, "certificatePassword");
            var reason = ArgumentHelper.GetString(arguments, "reason", "Document approval");
            var location = ArgumentHelper.GetString(arguments, "location", "");

            SecurityHelper.ValidateFilePath(certificatePath, "certificatePath", true);

            if (!File.Exists(certificatePath))
                throw new FileNotFoundException($"Certificate file not found: {certificatePath}");

            using var document = new Document(path);
            using var pdfSign = new PdfFileSignature(document);
            var pkcs = new PKCS7(certificatePath, certificatePassword)
            {
                Reason = reason,
                Location = location
            };

            var rect = new Rectangle(100, 100, 200, 100);
            pdfSign.Sign(1, true, rect, pkcs);
            pdfSign.Save(outputPath);
            return $"PDF digitally signed. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a signature from the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing signatureIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteSignature(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var signatureIndex = ArgumentHelper.GetInt(arguments, "signatureIndex");

            using var document = new Document(path);
            using var pdfSign = new PdfFileSignature(document);
            var signatureNames = pdfSign.GetSignNames();

            if (signatureIndex < 0 || signatureIndex >= signatureNames.Count)
                throw new ArgumentException($"signatureIndex must be between 0 and {signatureNames.Count - 1}");

            var signatureName = signatureNames[signatureIndex];
            pdfSign.RemoveSignature(signatureName);
            pdfSign.Save(outputPath);
            return $"Deleted signature '{signatureName}' (index {signatureIndex}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all signatures from the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <returns>JSON string with all signatures</returns>
    private Task<string> GetSignatures(string path)
    {
        return Task.Run(() =>
        {
            using var document = new Document(path);
            using var pdfSign = new PdfFileSignature(document);
            var signatureNames = pdfSign.GetSignNames();

            if (signatureNames.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    items = Array.Empty<object>(),
                    message = "No signatures found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var signatureList = new List<object>();
            for (var i = 0; i < signatureNames.Count; i++)
            {
                var signatureName = signatureNames[i];
                var signatureInfo = new Dictionary<string, object?>
                {
                    ["index"] = i,
                    ["name"] = signatureName
                };
                // Check signature validity - IsValid may not be available, use alternative method
                try
                {
                    _ = pdfSign.ExtractCertificate(signatureName);
                    signatureInfo["hasCertificate"] = true;
                }
                catch (Exception ex)
                {
                    signatureInfo["hasCertificate"] = false;
                    Console.Error.WriteLine(
                        $"[WARN] Failed to extract certificate for signature '{signatureName}': {ex.Message}");
                }

                signatureList.Add(signatureInfo);
            }

            var result = new
            {
                count = signatureList.Count,
                items = signatureList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}