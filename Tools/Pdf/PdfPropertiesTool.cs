using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing document properties in PDF files (get, set)
/// </summary>
public class PdfPropertiesTool : IAsposeTool
{
    public string Description => @"Manage document properties in PDF files. Supports 2 operations: get, set.

Usage examples:
- Get properties: pdf_properties(operation='get', path='doc.pdf')
- Set properties: pdf_properties(operation='set', path='doc.pdf', title='Title', author='Author', subject='Subject')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get': Get document properties (required params: path)
- 'set': Set document properties (required params: path)",
                @enum = new[] { "get", "set" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input for set)"
            },
            title = new
            {
                type = "string",
                description = "Title (for set)"
            },
            author = new
            {
                type = "string",
                description = "Author (for set)"
            },
            subject = new
            {
                type = "string",
                description = "Subject (for set)"
            },
            keywords = new
            {
                type = "string",
                description = "Keywords (for set)"
            },
            creator = new
            {
                type = "string",
                description = "Creator (for set)"
            },
            producer = new
            {
                type = "string",
                description = "Producer (for set)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        // Only get outputPath for operations that modify the document
        string? outputPath = null;
        if (operation.ToLower() == "set")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "get" => await GetProperties(path),
            "set" => await SetProperties(path, outputPath!, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets PDF properties
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <returns>JSON string with properties</returns>
    private Task<string> GetProperties(string path)
    {
        return Task.Run(() =>
        {
            using var document = new Document(path);
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
        });
    }

    /// <summary>
    ///     Sets PDF properties
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing various property values</param>
    /// <returns>Success message</returns>
    private Task<string> SetProperties(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var author = ArgumentHelper.GetStringNullable(arguments, "author");
            var subject = ArgumentHelper.GetStringNullable(arguments, "subject");
            var keywords = ArgumentHelper.GetStringNullable(arguments, "keywords");
            var creator = ArgumentHelper.GetStringNullable(arguments, "creator");
            var producer = ArgumentHelper.GetStringNullable(arguments, "producer");

            using var document = new Document(path);
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

            document.Save(outputPath);
            return $"Document properties updated. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets a PDF property with fallback to DocumentInfo
    /// </summary>
    /// <param name="document">PDF document</param>
    /// <param name="key">Metadata key name</param>
    /// <param name="value">Value to set</param>
    /// <param name="infoSetter">Fallback setter for DocumentInfo, null if no fallback</param>
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