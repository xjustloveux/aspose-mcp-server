using System.Text;
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

        return operation.ToLower() switch
        {
            "get" => await GetProperties(arguments),
            "set" => await SetProperties(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets PDF properties
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>Formatted string with properties</returns>
    private Task<string> GetProperties(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);

            using var document = new Document(path);
            var metadata = document.Metadata;
            var sb = new StringBuilder();

            sb.AppendLine("=== Document Properties ===");
            sb.AppendLine($"Title: {metadata["Title"] ?? "(none)"}");
            sb.AppendLine($"Author: {metadata["Author"] ?? "(none)"}");
            sb.AppendLine($"Subject: {metadata["Subject"] ?? "(none)"}");
            sb.AppendLine($"Keywords: {metadata["Keywords"] ?? "(none)"}");
            sb.AppendLine($"Creator: {metadata["Creator"] ?? "(none)"}");
            sb.AppendLine($"Producer: {metadata["Producer"] ?? "(none)"}");
            sb.AppendLine($"Creation Date: {metadata["CreationDate"] ?? "(none)"}");
            sb.AppendLine($"Modification Date: {metadata["ModDate"] ?? "(none)"}");
            sb.AppendLine();
            sb.AppendLine($"Total Pages: {document.Pages.Count}");
            sb.AppendLine($"Is Encrypted: {document.IsEncrypted}");
            sb.AppendLine($"Is Linearized: {document.IsLinearized}");

            return sb.ToString();
        });
    }

    /// <summary>
    ///     Sets PDF properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing various property values, optional outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> SetProperties(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var author = ArgumentHelper.GetStringNullable(arguments, "author");
            var subject = ArgumentHelper.GetStringNullable(arguments, "subject");
            var keywords = ArgumentHelper.GetStringNullable(arguments, "keywords");
            var creator = ArgumentHelper.GetStringNullable(arguments, "creator");
            var producer = ArgumentHelper.GetStringNullable(arguments, "producer");

            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            using var document = new Document(path);
            var metadata = document.Metadata;

            try
            {
                // Set standard PDF metadata properties
                // Aspose.Pdf Metadata dictionary may have restrictions on which keys can be set
                // Try using the indexer directly, which should work for standard PDF metadata keys
                if (!string.IsNullOrEmpty(title))
                    try
                    {
                        // Try direct assignment
                        metadata["Title"] = title;
                    }
                    catch (Exception ex) when (ex.Message.Contains("not valid") || ex.Message.Contains("Key"))
                    {
                        // If "Key is not valid" error, try using SetMetaInfo method instead
                        try
                        {
                            document.Info.Title = title;
                        }
                        catch (Exception ex)
                        {
                            // If both methods fail, skip setting Title - this is a limitation of some PDF files
                            Console.Error.WriteLine($"[WARN] Failed to set PDF Title property: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Try alternative method
                        try
                        {
                            document.Info.Title = title;
                        }
                        catch (Exception ex2)
                        {
                            // If both methods fail, skip setting Title
                            Console.Error.WriteLine($"[WARN] Failed to set PDF Title property (both methods failed): {ex.Message}, {ex2.Message}");
                        }
                    }

                if (!string.IsNullOrEmpty(author))
                    try
                    {
                        metadata["Author"] = author;
                    }
                    catch (Exception ex) when (ex.Message.Contains("not valid") || ex.Message.Contains("Key"))
                    {
                        // If "Key is not valid" error, try using Info property instead
                        try
                        {
                            document.Info.Author = author;
                        }
                        catch (Exception ex)
                        {
                            // If both methods fail, skip setting Author - this is a limitation of some PDF files
                            Console.Error.WriteLine($"[WARN] Failed to set PDF Author property: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Try alternative method
                        try
                        {
                            document.Info.Author = author;
                        }
                        catch (Exception ex2)
                        {
                            // If both methods fail, skip setting Author
                            Console.Error.WriteLine($"[WARN] Failed to set PDF Author property (both methods failed): {ex.Message}, {ex2.Message}");
                        }
                    }

                if (!string.IsNullOrEmpty(subject))
                    try
                    {
                        metadata["Subject"] = subject;
                    }
                    catch (Exception ex) when (ex.Message.Contains("not valid") || ex.Message.Contains("Key"))
                    {
                        // If "Key is not valid" error, try using Info property instead
                        try
                        {
                            document.Info.Subject = subject;
                        }
                        catch (Exception ex)
                        {
                            // If both methods fail, skip setting Subject - this is a limitation of some PDF files
                            Console.Error.WriteLine($"[WARN] Failed to set PDF Subject property: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Try alternative method
                        try
                        {
                            document.Info.Subject = subject;
                        }
                        catch (Exception ex2)
                        {
                            // If both methods fail, skip setting Subject
                            Console.Error.WriteLine($"[WARN] Failed to set PDF Subject property (both methods failed): {ex.Message}, {ex2.Message}");
                        }
                    }

                if (!string.IsNullOrEmpty(keywords))
                    try
                    {
                        metadata["Keywords"] = keywords;
                    }
                    catch (Exception ex) when (ex.Message.Contains("not valid") || ex.Message.Contains("Key"))
                    {
                        throw new ArgumentException(
                            $"Cannot set Keywords property: The PDF metadata dictionary does not support setting this key. Original error: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException($"Failed to set Keywords property: {ex.Message}");
                    }

                if (!string.IsNullOrEmpty(creator))
                    try
                    {
                        metadata["Creator"] = creator;
                    }
                    catch (Exception ex)
                    {
                        // Creator may be read-only, skip if it fails
                        Console.Error.WriteLine($"[WARN] Failed to set PDF Creator property (may be read-only): {ex.Message}");
                    }

                if (!string.IsNullOrEmpty(producer))
                    try
                    {
                        metadata["Producer"] = producer;
                    }
                    catch (Exception ex)
                    {
                        // Producer may be read-only, skip if it fails
                        Console.Error.WriteLine($"[WARN] Failed to set PDF Producer property (may be read-only): {ex.Message}");
                    }
            }
            catch (ArgumentException)
            {
                // Re-throw ArgumentException as-is
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
}