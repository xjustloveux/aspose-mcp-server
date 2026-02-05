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
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");

        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        DigitalSignatureUtil.RemoveAllSignatures(path, outputPath);

        return new SuccessResult
        {
            Message = "All digital signatures removed from the document."
        };
    }
}
