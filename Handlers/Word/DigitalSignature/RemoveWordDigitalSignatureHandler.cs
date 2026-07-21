using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.DigitalSignature;

/// <summary>
///     Handler for removing all digital signatures from a Word document.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemoveWordDigitalSignatureHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes all digital signatures from a Word document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (source file path), outputPath (destination file path)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when required parameters are missing or a path resolves outside the configured allowlist.
    /// </exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");

        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        // Both paths are filesystem sinks (read + write): resolve symlinks and enforce the
        // allowlist before either is touched.
        var allowedBasePaths = context.ServerConfig?.AllowedBasePaths ?? [];
        var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(path, allowedBasePaths, "path");
        var resolvedOutputPath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath, allowedBasePaths, "outputPath");

        var outputDir = Path.GetDirectoryName(resolvedOutputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        DigitalSignatureUtil.RemoveAllSignatures(resolvedPath, resolvedOutputPath);

        return new SuccessResult
        {
            Message = "All digital signatures removed from the document."
        };
    }
}
