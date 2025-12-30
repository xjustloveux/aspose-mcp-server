using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing attachments in PDF documents (add, delete, get)
/// </summary>
public class PdfAttachmentTool : IAsposeTool
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

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
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        string? outputPath = null;
        if (operation.ToLower() != "get")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddAttachment(path, outputPath!, arguments),
            "delete" => await DeleteAttachment(path, outputPath!, arguments),
            "get" => await GetAttachments(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an attachment to the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing attachmentPath, attachmentName, optional description</param>
    /// <returns>Success message</returns>
    /// <exception cref="FileNotFoundException">Thrown when attachment file does not exist</exception>
    /// <exception cref="ArgumentException">Thrown when attachment with same name already exists</exception>
    private Task<string> AddAttachment(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var attachmentPath = ArgumentHelper.GetString(arguments, "attachmentPath");
            var attachmentName = ArgumentHelper.GetString(arguments, "attachmentName");
            var description = ArgumentHelper.GetStringNullable(arguments, "description");

            SecurityHelper.ValidateFilePath(attachmentPath, "attachmentPath", true);
            SecurityHelper.ValidateStringLength(attachmentName, "attachmentName", 255);
            if (description != null)
                SecurityHelper.ValidateStringLength(description, "description", 1000);

            if (!File.Exists(attachmentPath))
                throw new FileNotFoundException($"Attachment file not found: {attachmentPath}");

            using var document = new Document(path);

            var existingNames = CollectAttachmentNames(document.EmbeddedFiles);
            if (existingNames.Any(n => string.Equals(n, attachmentName, StringComparison.OrdinalIgnoreCase) ||
                                       string.Equals(Path.GetFileName(n), attachmentName,
                                           StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException($"Attachment with name '{attachmentName}' already exists");

            var fileSpecification = new FileSpecification(attachmentPath, description ?? "")
            {
                Name = attachmentName
            };

            document.EmbeddedFiles.Add(fileSpecification);
            document.Save(outputPath);
            return $"Added attachment '{attachmentName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes an attachment from the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing attachmentName</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when attachment is not found</exception>
    private Task<string> DeleteAttachment(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var attachmentName = ArgumentHelper.GetString(arguments, "attachmentName");

            using var document = new Document(path);
            var embeddedFiles = document.EmbeddedFiles;

            var (found, actualName, attachmentNames) = FindAttachment(embeddedFiles, attachmentName);

            if (!found)
            {
                var availableNames = string.Join(", ", attachmentNames);
                throw new ArgumentException(
                    $"Attachment '{attachmentName}' not found. Available attachments: {(string.IsNullOrEmpty(availableNames) ? "(none)" : availableNames)}");
            }

            embeddedFiles.Delete(actualName);
            document.Save(outputPath);
            return $"Deleted attachment '{attachmentName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all attachments from the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <returns>JSON string with all attachments</returns>
    private Task<string> GetAttachments(string path)
    {
        return Task.Run(() =>
        {
            using var document = new Document(path);
            var embeddedFiles = document.EmbeddedFiles;

            if (embeddedFiles == null || embeddedFiles.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    items = Array.Empty<object>(),
                    message = "No attachments found"
                };
                return JsonSerializer.Serialize(emptyResult, JsonOptions);
            }

            var attachmentList = CollectAttachmentInfo(embeddedFiles);

            var result = new
            {
                count = attachmentList.Count,
                items = attachmentList
            };
            return JsonSerializer.Serialize(result, JsonOptions);
        });
    }

    /// <summary>
    ///     Collects all attachment names from the embedded files collection
    /// </summary>
    /// <param name="embeddedFiles">The embedded files collection</param>
    /// <returns>List of attachment names</returns>
    private static List<string> CollectAttachmentNames(EmbeddedFileCollection embeddedFiles)
    {
        var names = new List<string>();
        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                names.Add(file.Name ?? "");
            }
            catch
            {
                // Skip invalid entries
            }

        return names;
    }

    /// <summary>
    ///     Finds an attachment by name in the embedded files collection
    /// </summary>
    /// <param name="embeddedFiles">The embedded files collection</param>
    /// <param name="attachmentName">The name to search for</param>
    /// <returns>Tuple of (found, actualName, allNames)</returns>
    private static (bool found, string actualName, List<string> allNames) FindAttachment(
        EmbeddedFileCollection embeddedFiles, string attachmentName)
    {
        var allNames = new List<string>();

        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                var name = file.Name ?? "";
                allNames.Add(name);

                var fileName = Path.GetFileName(name);
                if (string.Equals(name, attachmentName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(fileName, attachmentName, StringComparison.OrdinalIgnoreCase))
                    return (true, name, allNames);
            }
            catch
            {
                // Skip invalid entries
            }

        return (false, "", allNames);
    }

    /// <summary>
    ///     Collects detailed attachment information from the embedded files collection
    /// </summary>
    /// <param name="embeddedFiles">The embedded files collection</param>
    /// <returns>List of attachment info objects</returns>
    private static List<object> CollectAttachmentInfo(EmbeddedFileCollection embeddedFiles)
    {
        var attachmentList = new List<object>();

        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                var attachmentInfo = new Dictionary<string, object?>
                {
                    ["index"] = i,
                    ["name"] = file.Name ?? "(unnamed)",
                    ["description"] = !string.IsNullOrEmpty(file.Description) ? file.Description : null,
                    ["mimeType"] = !string.IsNullOrEmpty(file.MIMEType) ? file.MIMEType : null
                };

                try
                {
                    if (file.Contents != null)
                        attachmentInfo["sizeBytes"] = file.Contents.Length;
                }
                catch
                {
                    attachmentInfo["sizeBytes"] = null;
                }

                attachmentList.Add(attachmentInfo);
            }
            catch (Exception ex)
            {
                attachmentList.Add(new { index = i, error = ex.Message });
            }

        return attachmentList;
    }
}