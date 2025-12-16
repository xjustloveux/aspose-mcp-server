using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

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
    /// Adds an attachment to the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, filePath, optional description, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddAttachment(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var attachmentPath = ArgumentHelper.GetString(arguments, "attachmentPath");
        var attachmentName = ArgumentHelper.GetString(arguments, "attachmentName");
        var description = ArgumentHelper.GetStringNullable(arguments, "description");

        // Validate paths
        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(attachmentPath, "attachmentPath");
        
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
        return await Task.FromResult($"Successfully added attachment '{attachmentName}'. Output: {outputPath}");
    }

    /// <summary>
    /// Deletes an attachment from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, attachmentName, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteAttachment(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var attachmentName = ArgumentHelper.GetString(arguments, "attachmentName");

        using var document = new Document(path);
        var embeddedFiles = document.EmbeddedFiles;
        
        for (int i = embeddedFiles.Count - 1; i >= 0; i--)
        {
            if (embeddedFiles[i].Name == attachmentName)
            {
                embeddedFiles.Delete(attachmentName);
                document.Save(outputPath);
                return await Task.FromResult($"Successfully deleted attachment '{attachmentName}'. Output: {outputPath}");
            }
        }

        throw new ArgumentException($"Attachment '{attachmentName}' not found");
    }

    /// <summary>
    /// Gets all attachments from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>Formatted string with all attachments</returns>
    private async Task<string> GetAttachments(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        using var document = new Document(path);
        var sb = new StringBuilder();
        sb.AppendLine("=== PDF Attachments ===");
        sb.AppendLine();

        var embeddedFiles = document.EmbeddedFiles;
        if (embeddedFiles.Count == 0)
        {
            sb.AppendLine("No attachments found.");
            return await Task.FromResult(sb.ToString());
        }

        sb.AppendLine($"Total Attachments: {embeddedFiles.Count}");
        sb.AppendLine();

        for (int i = 0; i < embeddedFiles.Count; i++)
        {
            var file = embeddedFiles[i];
            sb.AppendLine($"[{i}] Name: {file.Name}");
            if (!string.IsNullOrEmpty(file.Description))
                sb.AppendLine($"    Description: {file.Description}");
            if (file.Contents != null)
                sb.AppendLine($"    Size: {file.Contents.Length} bytes");
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

