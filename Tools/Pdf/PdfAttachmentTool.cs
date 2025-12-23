using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing attachments in PDF documents (add, delete, get)
/// </summary>
public class PdfAttachmentTool : IAsposeTool
{
    public string Description => @"Manage attachments in PDF documents. Supports 3 operations: add, delete, get.

Usage examples:
- Add attachment: pdf_attachment(operation='add', path='doc.pdf', attachmentPath='file.pdf', attachmentName='attachment.pdf')
- Delete attachment: pdf_attachment(operation='delete', path='doc.pdf', attachmentName='attachment.pdf')
- Get attachments: pdf_attachment(operation='get', path='doc.pdf')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add an attachment (required params: path, attachmentPath, attachmentName)
- 'delete': Delete an attachment (required params: path, attachmentName)
- 'get': Get all attachments (required params: path)",
                @enum = new[] { "add", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            attachmentPath = new
            {
                type = "string",
                description = "Attachment file path (required for add)"
            },
            attachmentName = new
            {
                type = "string",
                description = "Attachment name in PDF (required for add, delete)"
            },
            description = new
            {
                type = "string",
                description = "Attachment description (optional for add)"
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
            "add" => await AddAttachment(arguments),
            "delete" => await DeleteAttachment(arguments),
            "get" => await GetAttachments(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an attachment to the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, filePath, optional description, outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> AddAttachment(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var attachmentPath = ArgumentHelper.GetString(arguments, "attachmentPath");
            var attachmentName = ArgumentHelper.GetString(arguments, "attachmentName");
            var description = ArgumentHelper.GetStringNullable(arguments, "description");

            // Validate paths
            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
            SecurityHelper.ValidateFilePath(attachmentPath, "attachmentPath", true);

            // Validate attachment name length
            SecurityHelper.ValidateStringLength(attachmentName, "attachmentName", 255);
            if (description != null)
                SecurityHelper.ValidateStringLength(description, "description", 1000);

            if (!File.Exists(attachmentPath))
                throw new FileNotFoundException($"Attachment file not found: {attachmentPath}");

            using var document = new Document(path);
            var fileSpecification = new FileSpecification(attachmentPath, attachmentName);
            if (!string.IsNullOrEmpty(description))
                fileSpecification.Description = description;

            document.EmbeddedFiles.Add(fileSpecification);
            document.Save(outputPath);
            return $"Successfully added attachment '{attachmentName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes an attachment from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, attachmentName, optional outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteAttachment(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var attachmentName = ArgumentHelper.GetString(arguments, "attachmentName");

            using var document = new Document(path);
            var embeddedFiles = document.EmbeddedFiles;

            // Find and delete attachment by name
            // Note: EmbeddedFileCollection uses 1-based indexing for Item property
            var found = false;
            var attachmentNames = new List<string>();

            // First, collect all attachment names for debugging
            for (var i = 1; i <= embeddedFiles.Count; i++)
                try
                {
                    var file = embeddedFiles[i];
                    var name = file.Name ?? "";
                    attachmentNames.Add(name);

                    // Check Name property - use case-insensitive comparison
                    // Also check if the name ends with the attachment name (for full path cases)
                    var fileName = Path.GetFileName(name);
                    if (string.Equals(name, attachmentName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(fileName, attachmentName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Use the actual name from the file object for deletion
                        embeddedFiles.Delete(name);
                        found = true;
                        break;
                    }
                }
                catch (Exception ex)
                {
                    // Skip invalid indices
                    Console.Error.WriteLine($"[WARN] Error accessing attachment at index {i}: {ex.Message}");
                }

            if (!found)
            {
                var availableNames = string.Join(", ", attachmentNames);
                throw new ArgumentException(
                    $"Attachment '{attachmentName}' not found. Available attachments: {(string.IsNullOrEmpty(availableNames) ? "(none)" : availableNames)}");
            }

            document.Save(outputPath);
            return $"Successfully deleted attachment '{attachmentName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all attachments from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>Formatted string with all attachments</returns>
    private Task<string> GetAttachments(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);

            try
            {
                using var document = new Document(path);
                var sb = new StringBuilder();
                sb.AppendLine("=== PDF Attachments ===");
                sb.AppendLine();

                var embeddedFiles = document.EmbeddedFiles;
                if (embeddedFiles == null || embeddedFiles.Count == 0)
                {
                    sb.AppendLine("No attachments found.");
                    return sb.ToString();
                }

                sb.AppendLine($"Total Attachments: {embeddedFiles.Count}");
                sb.AppendLine();

                for (var i = 1; i <= embeddedFiles.Count; i++)
                    try
                    {
                        var file = embeddedFiles[i];
                        sb.AppendLine($"[{i}] Name: {file.Name ?? "(unnamed)"}");
                        if (!string.IsNullOrEmpty(file.Description))
                            sb.AppendLine($"    Description: {file.Description}");
                        try
                        {
                            if (file.Contents != null)
                                sb.AppendLine($"    Size: {file.Contents.Length} bytes");
                        }
                        catch (Exception ex)
                        {
                            sb.AppendLine("    Size: (unavailable)");
                            Console.Error.WriteLine($"[WARN] Failed to read attachment size: {ex.Message}");
                        }

                        sb.AppendLine();
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"[{i}] Error reading attachment: {ex.Message}");
                        sb.AppendLine();
                    }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Failed to get attachments: {ex.Message}");
            }
        });
    }
}