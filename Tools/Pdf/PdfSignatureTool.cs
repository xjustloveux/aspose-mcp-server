using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using Rectangle = System.Drawing.Rectangle;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing digital signatures in PDF documents (sign, delete, get)
/// </summary>
[McpServerToolType]
public class PdfSignatureTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfSignatureTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfSignatureTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_signature")]
    [Description(@"Manage digital signatures in PDF documents. Supports 3 operations: sign, delete, get.

Usage examples:
- Sign PDF: pdf_signature(operation='sign', path='doc.pdf', certificatePath='cert.pfx', certificatePassword='password')
- Sign with position: pdf_signature(operation='sign', path='doc.pdf', certificatePath='cert.pfx', certificatePassword='password', pageIndex=1, x=100, y=100, width=200, height=100)
- Sign with image: pdf_signature(operation='sign', path='doc.pdf', certificatePath='cert.pfx', certificatePassword='password', imagePath='stamp.png')
- Delete signature: pdf_signature(operation='delete', path='doc.pdf', signatureIndex=0)
- Get signatures: pdf_signature(operation='get', path='doc.pdf')")]
    public string Execute(
        [Description("Operation: sign, delete, get")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Path to certificate file (.pfx, required for sign)")]
        string? certificatePath = null,
        [Description("Certificate password (required for sign)")]
        string? certificatePassword = null,
        [Description("Reason for signing (for sign, optional)")]
        string reason = "Document approval",
        [Description("Location of signing (for sign, optional)")]
        string location = "",
        [Description("Signature index (0-based, required for delete)")]
        int signatureIndex = 0,
        [Description("Page index to place signature (1-based, for sign, default: 1)")]
        int pageIndex = 1,
        [Description("X position of signature in PDF coordinates (for sign, default: 100)")]
        int x = 100,
        [Description("Y position of signature in PDF coordinates (for sign, default: 100)")]
        int y = 100,
        [Description("Width of signature rectangle in PDF points (for sign, default: 200)")]
        int width = 200,
        [Description("Height of signature rectangle in PDF points (for sign, default: 100)")]
        int height = 100,
        [Description("Path to signature appearance image (for sign, optional)")]
        string? imagePath = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "sign" => SignDocument(ctx, outputPath, certificatePath, certificatePassword, reason, location, pageIndex,
                x, y, width, height, imagePath),
            "delete" => DeleteSignature(ctx, outputPath, signatureIndex),
            "get" => GetSignatures(ctx),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Digitally signs the PDF document using a certificate.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="certificatePath">The path to the certificate file (.pfx).</param>
    /// <param name="certificatePassword">The certificate password.</param>
    /// <param name="reason">The reason for signing.</param>
    /// <param name="location">The location of signing.</param>
    /// <param name="pageIndex">The 1-based page index to place the signature.</param>
    /// <param name="x">The X position of the signature.</param>
    /// <param name="y">The Y position of the signature.</param>
    /// <param name="width">The width of the signature rectangle.</param>
    /// <param name="height">The height of the signature rectangle.</param>
    /// <param name="imagePath">Optional path to a signature appearance image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when certificate or image files are not found.</exception>
    private static string SignDocument(DocumentContext<Document> ctx, string? outputPath,
        string? certificatePath, string? certificatePassword,
        string reason, string location,
        int pageIndex, int x, int y, int width, int height, string? imagePath)
    {
        if (string.IsNullOrEmpty(certificatePath))
            throw new ArgumentException("certificatePath is required for sign operation");
        if (string.IsNullOrEmpty(certificatePassword))
            throw new ArgumentException("certificatePassword is required for sign operation");

        SecurityHelper.ValidateFilePath(certificatePath, "certificatePath", true);

        if (!File.Exists(certificatePath))
            throw new FileNotFoundException($"Certificate file not found: {certificatePath}");

        var document = ctx.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        using var pdfSign = new PdfFileSignature(document);
        var pkcs = new PKCS7(certificatePath, certificatePassword)
        {
            Reason = reason,
            Location = location
        };

        if (!string.IsNullOrEmpty(imagePath))
        {
            SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Signature image file not found: {imagePath}");
            pdfSign.SignatureAppearance = imagePath;
        }

        var rect = new Rectangle(x, y, width, height);
        pdfSign.Sign(pageIndex, true, rect, pkcs);

        ctx.Save(outputPath);

        return $"PDF digitally signed on page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a digital signature from the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="signatureIndex">The 0-based signature index to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the signature index is invalid.</exception>
    private static string DeleteSignature(DocumentContext<Document> ctx, string? outputPath, int signatureIndex)
    {
        var document = ctx.Document;
        using var pdfSign = new PdfFileSignature(document);
        var signatureNames = pdfSign.GetSignNames();

        if (signatureIndex < 0 || signatureIndex >= signatureNames.Count)
            throw new ArgumentException($"signatureIndex must be between 0 and {signatureNames.Count - 1}");

        var signatureName = signatureNames[signatureIndex];
        pdfSign.RemoveSignature(signatureName);

        ctx.Save(outputPath);

        return $"Deleted signature '{signatureName}' (index {signatureIndex}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Retrieves all digital signatures from the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing signature information.</returns>
    private static string GetSignatures(DocumentContext<Document> ctx)
    {
        var document = ctx.Document;
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

        List<object> signatureList = [];
        for (var i = 0; i < signatureNames.Count; i++)
        {
            var signatureName = signatureNames[i];
            var signatureInfo = new Dictionary<string, object?>
            {
                ["index"] = i,
                ["name"] = signatureName
            };

            try
            {
                signatureInfo["isValid"] = pdfSign.VerifySignature(signatureName);
            }
            catch
            {
                signatureInfo["isValid"] = false;
            }

            try
            {
                _ = pdfSign.ExtractCertificate(signatureName);
                signatureInfo["hasCertificate"] = true;
            }
            catch
            {
                signatureInfo["hasCertificate"] = false;
            }

            signatureList.Add(signatureInfo);
        }

        var result = new
        {
            count = signatureList.Count,
            items = signatureList
        };
        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}