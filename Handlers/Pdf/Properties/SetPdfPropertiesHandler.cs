using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Properties;

/// <summary>
///     Handler for setting document properties in PDF files.
/// </summary>
public class SetPdfPropertiesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set";

    /// <summary>
    ///     Sets document properties in the PDF.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: title, author, subject, keywords, creator, producer
    /// </param>
    /// <returns>Success message indicating properties were updated.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var title = parameters.GetOptional<string?>("title");
        var author = parameters.GetOptional<string?>("author");
        var subject = parameters.GetOptional<string?>("subject");
        var keywords = parameters.GetOptional<string?>("keywords");
        var creator = parameters.GetOptional<string?>("creator");
        var producer = parameters.GetOptional<string?>("producer");

        var document = context.Document;
        var docInfo = document.Info;

        try
        {
            SetPropertyWithFallback(document, "Title", title, v => docInfo.Title = v);
            SetPropertyWithFallback(document, "Author", author, v => docInfo.Author = v);
            SetPropertyWithFallback(document, "Subject", subject, v => docInfo.Subject = v);
            SetPropertyWithFallback(document, "Keywords", keywords, v => docInfo.Keywords = v);
            SetPropertyWithFallback(document, "Creator", creator, null);
            SetPropertyWithFallback(document, "Producer", producer, null);
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Failed to set document properties: {ex.Message}. Note: Some PDF files may have restrictions on modifying metadata, or the document may be encrypted/protected.");
        }

        MarkModified(context);

        return Success("Document properties updated.");
    }

    /// <summary>
    ///     Sets a document property with fallback to DocumentInfo if metadata fails.
    /// </summary>
    /// <param name="document">The PDF document.</param>
    /// <param name="key">The property key to set.</param>
    /// <param name="value">The value to set.</param>
    /// <param name="infoSetter">Optional fallback setter using DocumentInfo.</param>
    private static void SetPropertyWithFallback(Document document, string key, string? value,
        Action<string>? infoSetter)
    {
        if (string.IsNullOrEmpty(value)) return;

        try
        {
            document.Metadata[key] = value;
        }
        catch
        {
            if (infoSetter != null)
                try
                {
                    infoSetter(value);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[WARN] Failed to set PDF {key} property: {ex.Message}");
                }
            else
                Console.Error.WriteLine($"[WARN] Failed to set PDF {key} property (may be read-only)");
        }
    }
}
