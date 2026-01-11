using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Signature;

/// <summary>
///     Handler for signing PDF documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var certificatePath = parameters.GetRequired<string>("certificatePath");
        var password = parameters.GetRequired<string>("password");
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var reason = parameters.GetOptional("reason", "Document Signed");
        var location = parameters.GetOptional("location", "");
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 100.0);
        var width = parameters.GetOptional("width", 200.0);
        var height = parameters.GetOptional("height", 100.0);

        SecurityHelper.ValidateFilePath(certificatePath, "certificatePath", true);

        if (!File.Exists(certificatePath))
            throw new FileNotFoundException($"Certificate file not found: {certificatePath}", certificatePath);

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var signatureField = new SignatureField(document.Pages[pageIndex],
            new Rectangle(x, y, x + width, y + height))
        {
            Name = $"Signature_{DateTime.Now:yyyyMMddHHmmss}"
        };

        document.Form.Add(signatureField);

        var pkcs = new PKCS7(certificatePath, password)
        {
            Reason = reason,
            Location = location,
            Date = DateTime.Now
        };

        signatureField.Sign(pkcs);

        MarkModified(context);

        return Success($"Document signed on page {pageIndex}.");
    }
}
