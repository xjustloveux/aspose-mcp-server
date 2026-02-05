using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.DigitalSignature;

/// <summary>
///     Handler for signing a Word document with a digital signature.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SignWordDigitalSignatureHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "sign";

    /// <summary>
    ///     Signs a Word document with a digital certificate.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (source file path), outputPath (destination file path),
    ///     certificatePath (PFX certificate file path), certificatePassword (certificate password)
    ///     Optional: comments (signature comments)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the certificate file is not found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSignParameters(parameters);

        SecurityHelper.ValidateFilePath(p.Path, allowAbsolutePaths: true);
        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);
        SecurityHelper.ValidateFilePath(p.CertificatePath, "certificatePath", true);

        if (!System.IO.File.Exists(p.CertificatePath))
            throw new FileNotFoundException($"Certificate file not found: {p.CertificatePath}");

        var outputDir = Path.GetDirectoryName(p.OutputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var certificateHolder = CertificateHolder.Create(p.CertificatePath, p.CertificatePassword);

        var signOptions = new SignOptions
        {
            Comments = p.Comments ?? string.Empty,
            SignTime = DateTime.Now
        };

        DigitalSignatureUtil.Sign(p.Path, p.OutputPath, certificateHolder, signOptions);

        return new SuccessResult
        {
            Message = "Document signed with digital signature successfully."
        };
    }

    /// <summary>
    ///     Extracts sign parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted sign parameters.</returns>
    private static SignParameters ExtractSignParameters(OperationParameters parameters)
    {
        return new SignParameters(
            parameters.GetRequired<string>("path"),
            parameters.GetRequired<string>("outputPath"),
            parameters.GetRequired<string>("certificatePath"),
            parameters.GetRequired<string>("certificatePassword"),
            parameters.GetOptional<string?>("comments")
        );
    }

    /// <summary>
    ///     Parameters for the sign operation.
    /// </summary>
    /// <param name="Path">The source document file path.</param>
    /// <param name="OutputPath">The destination file path for the signed document.</param>
    /// <param name="CertificatePath">The PFX certificate file path.</param>
    /// <param name="CertificatePassword">The certificate password.</param>
    /// <param name="Comments">Optional comments for the signature.</param>
    private sealed record SignParameters(
        string Path,
        string OutputPath,
        string CertificatePath,
        string CertificatePassword,
        string? Comments);
}
