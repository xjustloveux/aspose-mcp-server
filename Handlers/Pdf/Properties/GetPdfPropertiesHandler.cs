using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Properties;

/// <summary>
///     Handler for retrieving document properties from PDF files.
/// </summary>
public class GetPdfPropertiesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets document properties from the PDF.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No additional parameters required.</param>
    /// <returns>JSON string containing document properties.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        var metadata = document.Metadata;

        var result = new
        {
            title = metadata["Title"]?.ToString(),
            author = metadata["Author"]?.ToString(),
            subject = metadata["Subject"]?.ToString(),
            keywords = metadata["Keywords"]?.ToString(),
            creator = metadata["Creator"]?.ToString(),
            producer = metadata["Producer"]?.ToString(),
            creationDate = metadata["CreationDate"]?.ToString(),
            modificationDate = metadata["ModDate"]?.ToString(),
            totalPages = document.Pages.Count,
            isEncrypted = document.IsEncrypted,
            isLinearized = document.IsLinearized
        };

        return JsonResult(result);
    }
}
