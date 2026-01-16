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
        var createParams = ExtractCreateParameters(parameters);

        SecurityHelper.ValidateFilePath(createParams.OutputPath, "outputPath", true);

        using var document = new Document();
        document.Pages.Add();
        document.Save(createParams.OutputPath);

        return Success($"PDF document created. Output: {createParams.OutputPath}");
    }

    /// <summary>
    ///     Extracts create parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted create parameters.</returns>
    private static CreateParameters ExtractCreateParameters(OperationParameters parameters)
    {
        return new CreateParameters(
            parameters.GetRequired<string>("outputPath")
        );
    }

    /// <summary>
    ///     Record to hold create parameters.
    /// </summary>
    private sealed record CreateParameters(string OutputPath);
}
