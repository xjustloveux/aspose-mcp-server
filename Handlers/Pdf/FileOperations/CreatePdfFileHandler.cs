using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for creating a new empty PDF document.
/// </summary>
public class CreatePdfFileHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new empty PDF document with one blank page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputPath
    /// </param>
    /// <returns>Success message with output path.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetRequired<string>("outputPath");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        using var document = new Document();
        document.Pages.Add();
        document.Save(outputPath);

        return Success($"PDF document created. Output: {outputPath}");
    }
}
