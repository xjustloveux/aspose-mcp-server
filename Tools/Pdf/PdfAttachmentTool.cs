using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing attachments in PDF documents (add, delete, get)
/// </summary>
[McpServerToolType]
public class PdfAttachmentTool
{
    /// <summary>
    ///     JSON serialization options for formatted output.
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfAttachmentTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfAttachmentTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_attachment")]
    [Description(@"Manage attachments in PDF documents. Supports 3 operations: add, delete, get.

Usage examples:
- Add attachment: pdf_attachment(operation='add', path='doc.pdf', attachmentPath='file.pdf', attachmentName='attachment.pdf')
- Delete attachment: pdf_attachment(operation='delete', path='doc.pdf', attachmentName='attachment.pdf')
- Get attachments: pdf_attachment(operation='get', path='doc.pdf')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add an attachment (required params: path, attachmentPath, attachmentName)
- 'delete': Delete an attachment (required params: path, attachmentName)
- 'get': Get all attachments (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Attachment file path (required for add)")]
        string? attachmentPath = null,
        [Description("Attachment name in PDF (required for add, delete)")]
        string? attachmentName = null,
        [Description("Attachment description (optional for add)")]
        string? description = null)
    {
        return operation.ToLower() switch
        {
            "add" => AddAttachment(sessionId, path, outputPath, attachmentPath, attachmentName, description),
            "delete" => DeleteAttachment(sessionId, path, outputPath, attachmentName),
            "get" => GetAttachments(sessionId, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a file attachment to the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="attachmentPath">The path to the file to attach.</param>
    /// <param name="attachmentName">The name for the attachment in the PDF.</param>
    /// <param name="description">Optional description for the attachment.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or attachment already exists.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the attachment file is not found.</exception>
    private string AddAttachment(string? sessionId, string? path, string? outputPath, string? attachmentPath,
        string? attachmentName, string? description)
    {
        if (string.IsNullOrEmpty(attachmentPath))
            throw new ArgumentException("attachmentPath is required for add operation");
        if (string.IsNullOrEmpty(attachmentName))
            throw new ArgumentException("attachmentName is required for add operation");

        SecurityHelper.ValidateFilePath(attachmentPath, "attachmentPath", true);
        SecurityHelper.ValidateStringLength(attachmentName, "attachmentName", 255);
        if (description != null)
            SecurityHelper.ValidateStringLength(description, "description", 1000);

        if (!File.Exists(attachmentPath))
            throw new FileNotFoundException($"Attachment file not found: {attachmentPath}");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

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
        ctx.Save(outputPath);
        return $"Added attachment '{attachmentName}'. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes an attachment from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="attachmentName">The name of the attachment to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the attachment is not found.</exception>
    private string DeleteAttachment(string? sessionId, string? path, string? outputPath, string? attachmentName)
    {
        if (string.IsNullOrEmpty(attachmentName))
            throw new ArgumentException("attachmentName is required for delete operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;
        var embeddedFiles = document.EmbeddedFiles;

        var (found, actualName, attachmentNames) = FindAttachment(embeddedFiles, attachmentName);

        if (!found)
        {
            var availableNames = string.Join(", ", attachmentNames);
            throw new ArgumentException(
                $"Attachment '{attachmentName}' not found. Available attachments: {(string.IsNullOrEmpty(availableNames) ? "(none)" : availableNames)}");
        }

        embeddedFiles.Delete(actualName);
        ctx.Save(outputPath);
        return $"Deleted attachment '{attachmentName}'. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Retrieves all attachments from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <returns>A JSON string containing attachment information.</returns>
    private string GetAttachments(string? sessionId, string? path)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;
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
    }

    /// <summary>
    ///     Collects all attachment names from the embedded files collection.
    /// </summary>
    /// <param name="embeddedFiles">The collection of embedded files.</param>
    /// <returns>A list of attachment names.</returns>
    private static List<string> CollectAttachmentNames(EmbeddedFileCollection embeddedFiles)
    {
        List<string> names = [];
        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                names.Add(file.Name ?? "");
            }
            catch
            {
                // Ignore errors reading individual attachment names
            }

        return names;
    }

    /// <summary>
    ///     Finds an attachment by name in the embedded files collection.
    /// </summary>
    /// <param name="embeddedFiles">The collection of embedded files.</param>
    /// <param name="attachmentName">The name of the attachment to find.</param>
    /// <returns>A tuple indicating whether the attachment was found, its actual name, and all available names.</returns>
    private static (bool found, string actualName, List<string> allNames) FindAttachment(
        EmbeddedFileCollection embeddedFiles, string attachmentName)
    {
        List<string> allNames = [];

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
                // Ignore errors reading individual attachment
            }

        return (false, "", allNames);
    }

    /// <summary>
    ///     Collects detailed information about all attachments.
    /// </summary>
    /// <param name="embeddedFiles">The collection of embedded files.</param>
    /// <returns>A list of attachment information objects.</returns>
    private static List<object> CollectAttachmentInfo(EmbeddedFileCollection embeddedFiles)
    {
        List<object> attachmentList = [];

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