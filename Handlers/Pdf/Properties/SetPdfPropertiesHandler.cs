using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Properties;

/// <summary>
///     Handler for setting document properties in PDF files.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetParameters(parameters);

        var document = context.Document;
        var docInfo = document.Info;

        try
        {
            SetPropertyWithFallback(document, "Title", p.Title, v => docInfo.Title = v);
            SetPropertyWithFallback(document, "Author", p.Author, v => docInfo.Author = v);
            SetPropertyWithFallback(document, "Subject", p.Subject, v => docInfo.Subject = v);
            SetPropertyWithFallback(document, "Keywords", p.Keywords, v => docInfo.Keywords = v);
            SetPropertyWithFallback(document, "Creator", p.Creator, null);
            SetPropertyWithFallback(document, "Producer", p.Producer, null);
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

        return new SuccessResult { Message = "Document properties updated." };
    }

    /// <summary>
    ///     Extracts parameters for set operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetParameters ExtractSetParameters(OperationParameters parameters)
    {
        return new SetParameters(
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("author"),
            parameters.GetOptional<string?>("subject"),
            parameters.GetOptional<string?>("keywords"),
            parameters.GetOptional<string?>("creator"),
            parameters.GetOptional<string?>("producer")
        );
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

    /// <summary>
    ///     Parameters for set operation.
    /// </summary>
    /// <param name="Title">The document title.</param>
    /// <param name="Author">The document author.</param>
    /// <param name="Subject">The document subject.</param>
    /// <param name="Keywords">The document keywords.</param>
    /// <param name="Creator">The document creator.</param>
    /// <param name="Producer">The document producer.</param>
    private sealed record SetParameters(
        string? Title,
        string? Author,
        string? Subject,
        string? Keywords,
        string? Creator,
        string? Producer);
}
