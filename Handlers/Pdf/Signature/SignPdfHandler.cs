using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Signature;

/// <summary>
///     Handler for signing PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SignPdfHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "sign";

    /// <summary>
    ///     Signs a PDF document with a digital signature.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: certificatePath, password.
    ///     Optional: pageIndex (default: 1), reason, location, x, y, width, height.
    /// </param>
    /// <returns>Success message with signature details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSignParameters(parameters);

        SecurityHelper.ValidateFilePath(p.CertificatePath, "certificatePath", true);

        if (!File.Exists(p.CertificatePath))
            throw new FileNotFoundException($"Certificate file not found: {p.CertificatePath}", p.CertificatePath);

        var document = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var signatureField = new SignatureField(document.Pages[p.PageIndex],
            new Rectangle(p.X, p.Y, p.X + p.Width, p.Y + p.Height))
        {
            Name = $"Signature_{DateTime.Now:yyyyMMddHHmmss}"
        };

        document.Form.Add(signatureField);

        var pkcs = new PKCS7(p.CertificatePath, p.Password)
        {
            Reason = p.Reason,
            Location = p.Location,
            Date = DateTime.Now
        };

        signatureField.Sign(pkcs);

        MarkModified(context);

        return new SuccessResult { Message = $"Document signed on page {p.PageIndex}." };
    }

    /// <summary>
    ///     Extracts parameters for sign operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SignParameters ExtractSignParameters(OperationParameters parameters)
    {
        return new SignParameters(
            parameters.GetRequired<string>("certificatePath"),
            parameters.GetRequired<string>("password"),
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("reason", "Document Signed"),
            parameters.GetOptional("location", ""),
            parameters.GetOptional("x", 100.0),
            parameters.GetOptional("y", 100.0),
            parameters.GetOptional("width", 200.0),
            parameters.GetOptional("height", 100.0)
        );
    }

    /// <summary>
    ///     Parameters for sign operation.
    /// </summary>
    /// <param name="CertificatePath">The path to the certificate file.</param>
    /// <param name="Password">The certificate password.</param>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="Reason">The signing reason.</param>
    /// <param name="Location">The signing location.</param>
    /// <param name="X">The X coordinate of the signature.</param>
    /// <param name="Y">The Y coordinate of the signature.</param>
    /// <param name="Width">The width of the signature.</param>
    /// <param name="Height">The height of the signature.</param>
    private sealed record SignParameters(
        string CertificatePath,
        string Password,
        int PageIndex,
        string Reason,
        string Location,
        double X,
        double Y,
        double Width,
        double Height);
}
