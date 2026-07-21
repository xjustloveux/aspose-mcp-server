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
    /// <exception cref="ArgumentException">
    ///     Thrown when required parameters are missing or a path resolves outside the configured allowlist.
    /// </exception>
    /// <exception cref="FileNotFoundException">Thrown when the certificate file is not found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSignParameters(parameters);

        SecurityHelper.ValidateFilePath(p.Path, allowAbsolutePaths: true);
        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);
        SecurityHelper.ValidateFilePath(p.CertificatePath, "certificatePath", true);

        // All three paths are filesystem sinks (two reads, one write): resolve symlinks and
        // enforce the allowlist before any of them is touched.
        var allowedBasePaths = context.ServerConfig?.AllowedBasePaths ?? [];
        var sourcePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.Path, allowedBasePaths, "path");
        var outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath, allowedBasePaths, "outputPath");
        var certificatePath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(p.CertificatePath, allowedBasePaths, "certificatePath");

        if (!System.IO.File.Exists(certificatePath))
            throw new FileNotFoundException("The specified file was not found.");

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var certificateHolder = CertificateHolder.Create(certificatePath, p.CertificatePassword);

        var signOptions = new SignOptions
        {
            Comments = p.Comments ?? string.Empty,
            SignTime = DateTime.Now
        };

        DigitalSignatureUtil.Sign(sourcePath, outputPath, certificateHolder, signOptions);

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
