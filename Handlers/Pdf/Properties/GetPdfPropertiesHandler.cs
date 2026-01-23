using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Properties;

namespace AsposeMcpServer.Handlers.Pdf.Properties;

/// <summary>
///     Handler for retrieving document properties from PDF files.
/// </summary>
[ResultType(typeof(GetPropertiesPdfResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        var metadata = document.Metadata;
        var info = document.Info;

        var result = new GetPropertiesPdfResult
        {
            Title = GetPropertyValue(metadata, "Title") ?? info.Title,
            Author = GetPropertyValue(metadata, "Author") ?? info.Author,
            Subject = GetPropertyValue(metadata, "Subject") ?? info.Subject,
            Keywords = GetPropertyValue(metadata, "Keywords") ?? info.Keywords,
            Creator = GetPropertyValue(metadata, "Creator") ?? info.Creator,
            Producer = GetPropertyValue(metadata, "Producer") ?? info.Producer,
            CreationDate = GetPropertyValue(metadata, "CreationDate") ?? FormatDate(info.CreationDate),
            ModificationDate = GetPropertyValue(metadata, "ModDate") ?? FormatDate(info.ModDate),
            TotalPages = document.Pages.Count,
            IsEncrypted = document.IsEncrypted,
            IsLinearized = document.IsLinearized
        };

        return result;
    }

    /// <summary>
    ///     Gets a property value from metadata, returning null if not found or on error.
    /// </summary>
    /// <param name="metadata">The document metadata.</param>
    /// <param name="key">The property key.</param>
    /// <returns>The property value as string, or null.</returns>
    private static string? GetPropertyValue(Metadata metadata, string key)
    {
        try
        {
            return metadata[key]?.ToString();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    ///     Formats a DateTime to string, returning null for default/minimum values.
    /// </summary>
    /// <param name="date">The date to format.</param>
    /// <returns>The formatted date string, or null.</returns>
    private static string? FormatDate(DateTime date)
    {
        if (date == default || date == DateTime.MinValue)
            return null;
        return date.ToString("O");
    }
}
