using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing document properties in PDF files (get, set)
/// </summary>
[McpServerToolType]
public class PdfPropertiesTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfPropertiesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfPropertiesTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_properties")]
    [Description(@"Manage document properties in PDF files. Supports 2 operations: get, set.

Usage examples:
- Get properties: pdf_properties(operation='get', path='doc.pdf')
- Set properties: pdf_properties(operation='set', path='doc.pdf', title='Title', author='Author', subject='Subject')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get': Get document properties (required params: path)
- 'set': Set document properties (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input for set)")]
        string? outputPath = null,
        [Description("Title (for set)")] string? title = null,
        [Description("Author (for set)")] string? author = null,
        [Description("Subject (for set)")] string? subject = null,
        [Description("Keywords (for set)")] string? keywords = null,
        [Description("Creator (for set)")] string? creator = null,
        [Description("Producer (for set)")] string? producer = null)
    {
        return operation.ToLower() switch
        {
            "get" => GetProperties(sessionId, path),
            "set" => SetProperties(sessionId, path, outputPath, title, author, subject, keywords, creator, producer),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Retrieves document properties from the PDF.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <returns>A JSON string containing document properties.</returns>
    private string GetProperties(string? sessionId, string? path)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;
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

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Sets document properties in the PDF.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="title">Optional document title.</param>
    /// <param name="author">Optional document author.</param>
    /// <param name="subject">Optional document subject.</param>
    /// <param name="keywords">Optional document keywords.</param>
    /// <param name="creator">Optional document creator.</param>
    /// <param name="producer">Optional document producer.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when property modification fails.</exception>
    private string SetProperties(string? sessionId, string? path, string? outputPath, string? title, string? author,
        string? subject, string? keywords, string? creator, string? producer)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;
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

        ctx.Save(outputPath);

        return $"Document properties updated. {ctx.GetOutputMessage(outputPath)}";
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