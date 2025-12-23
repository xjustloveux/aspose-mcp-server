using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document protection (protect, unprotect)
///     Merges: WordProtectTool, WordUnprotectTool
/// </summary>
public class WordProtectionTool : IAsposeTool
{
    public string Description => @"Protect or unprotect a Word document. Supports 2 operations: protect, unprotect.

Usage examples:
- Protect document: word_protection(operation='protect', path='doc.docx', password='password', protectionType='ReadOnly')
- Unprotect document: word_protection(operation='unprotect', path='doc.docx', password='password')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'protect': Protect document (required params: path, password, protectionType)
- 'unprotect': Unprotect document (required params: path, password)",
                @enum = new[] { "protect", "unprotect" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            password = new
            {
                type = "string",
                description = "Protection password (required for protect operation, optional for unprotect)"
            },
            protectionType = new
            {
                type = "string",
                description =
                    "Protection type: 'ReadOnly', 'AllowOnlyComments', 'AllowOnlyFormFields', 'AllowOnlyRevisions' (required for protect operation)",
                @enum = new[] { "ReadOnly", "AllowOnlyComments", "AllowOnlyFormFields", "AllowOnlyRevisions" }
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

        return operation.ToLower() switch
        {
            "protect" => await ProtectAsync(arguments, path),
            "unprotect" => await UnprotectAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Protects the document with password
    /// </summary>
    /// <param name="arguments">JSON arguments containing password, optional protectionType, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private Task<string> ProtectAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var password = ArgumentHelper.GetString(arguments, "password");
            var protectionTypeStr = ArgumentHelper.GetString(arguments, "protectionType", "ReadOnly");

            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);

            var protectionType = protectionTypeStr switch
            {
                "ReadOnly" => ProtectionType.ReadOnly,
                "AllowOnlyComments" => ProtectionType.AllowOnlyComments,
                "AllowOnlyFormFields" => ProtectionType.AllowOnlyFormFields,
                "AllowOnlyRevisions" => ProtectionType.AllowOnlyRevisions,
                _ => ProtectionType.ReadOnly
            };

            doc.Protect(protectionType, password);
            doc.Save(outputPath);

            return $"Document protected with {protectionType}: {outputPath}";
        });
    }

    /// <summary>
    ///     Removes protection from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing password, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private Task<string> UnprotectAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var password = ArgumentHelper.GetStringNullable(arguments, "password");

            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);
            var wasProtected = doc.ProtectionType != ProtectionType.NoProtection;

            if (!wasProtected)
            {
                if (!string.Equals(path, outputPath, StringComparison.OrdinalIgnoreCase))
                {
                    doc.Save(outputPath);
                    return $"Document is not protected, saved to: {outputPath}";
                }

                return "Document is not protected, no need to unprotect";
            }

            doc.Unprotect(password);

            if (doc.ProtectionType != ProtectionType.NoProtection)
                throw new InvalidOperationException(
                    "Unprotect failed, password may be incorrect or document is restricted");

            doc.Save(outputPath);
            return $"Protection removed successfully\nOutput: {outputPath}";
        });
    }
}