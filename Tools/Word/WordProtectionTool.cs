using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Loading;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document protection (protect, unprotect)
///     Merges: WordProtectTool, WordUnprotectTool
/// </summary>
public class WordProtectionTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Protect or unprotect a Word document. Supports 2 operations: protect, unprotect.

Usage examples:
- Protect document: word_protection(operation='protect', path='doc.docx', password='password', protectionType='ReadOnly')
- Unprotect document: word_protection(operation='unprotect', path='doc.docx', password='password')

Protection types:
- ReadOnly: Prevent all modifications (most restrictive)
- AllowOnlyComments: Allow only adding comments
- AllowOnlyFormFields: Allow only filling in form fields
- AllowOnlyRevisions: Allow only tracked changes

Notes:
- Password is required for 'protect' operation (cannot be empty)
- Password is optional for 'unprotect' (some documents may not require password)
- If unprotect fails, verify the password is correct
- For encrypted documents (with open password), the same password will be used to open the file";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "protect" => await ProtectAsync(path, outputPath, arguments),
            "unprotect" => await UnprotectAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Protects the document with specified protection type and password
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing password and protectionType</param>
    /// <returns>Success message with protection type and output path</returns>
    private Task<string> ProtectAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var password = ArgumentHelper.GetStringNullable(arguments, "password");
            var protectionTypeStr = ArgumentHelper.GetString(arguments, "protectionType", "ReadOnly");

            if (string.IsNullOrWhiteSpace(password))
                throw new ArgumentException(
                    "Password is required for protect operation. Please provide a non-empty password.");

            var doc = LoadDocument(path, password);
            var protectionType = GetProtectionType(protectionTypeStr);

            doc.Protect(protectionType, password);
            doc.Save(outputPath);

            return $"Document protected with {protectionType}: {outputPath}";
        });
    }

    /// <summary>
    ///     Loads a Word document with optional password for encrypted files
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="password">Optional password for encrypted documents</param>
    /// <returns>Loaded Document object</returns>
    /// <exception cref="InvalidOperationException">Thrown when document cannot be loaded</exception>
    private static Document LoadDocument(string path, string? password)
    {
        try
        {
            // First try with password if provided (for encrypted documents)
            if (!string.IsNullOrEmpty(password))
                try
                {
                    var loadOptions = new LoadOptions { Password = password };
                    return new Document(path, loadOptions);
                }
                catch (IncorrectPasswordException)
                {
                    // Password didn't work for opening, try without password
                    // (file might not be encrypted, password is for protection)
                }

            // Try loading without password
            return new Document(path);
        }
        catch (FileNotFoundException)
        {
            throw new InvalidOperationException($"Document not found: {path}");
        }
        catch (IncorrectPasswordException)
        {
            throw new InvalidOperationException(
                $"Document is encrypted and requires a password to open: {path}");
        }
        catch (UnsupportedFileFormatException ex)
        {
            throw new InvalidOperationException(
                $"Unsupported file format or corrupted document: {path}. Details: {ex.Message}");
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException(
                $"Failed to load document: {path}. Details: {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     Converts a protection type string to ProtectionType enum using Enum.TryParse
    /// </summary>
    /// <param name="protectionTypeStr">
    ///     Protection type string: ReadOnly, AllowOnlyComments, AllowOnlyFormFields,
    ///     AllowOnlyRevisions
    /// </param>
    /// <returns>Corresponding ProtectionType enum value, defaults to ReadOnly</returns>
    private static ProtectionType GetProtectionType(string protectionTypeStr)
    {
        if (Enum.TryParse<ProtectionType>(protectionTypeStr, true, out var result))
            return result;

        return ProtectionType.ReadOnly;
    }

    /// <summary>
    ///     Removes protection from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional password</param>
    /// <returns>Success message with output path, or error message if unprotect failed</returns>
    private Task<string> UnprotectAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var password = ArgumentHelper.GetStringNullable(arguments, "password");

            var doc = LoadDocument(path, password);
            var previousProtectionType = doc.ProtectionType;

            if (previousProtectionType == ProtectionType.NoProtection)
            {
                if (!string.Equals(path, outputPath, StringComparison.OrdinalIgnoreCase))
                {
                    doc.Save(outputPath);
                    return $"Document is not protected, saved to: {outputPath}";
                }

                return "Document is not protected, no need to unprotect";
            }

            try
            {
                doc.Unprotect(password);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to unprotect document: The password may be incorrect or the document has additional restrictions. Details: {ex.Message}",
                    ex);
            }

            if (doc.ProtectionType != ProtectionType.NoProtection)
                throw new InvalidOperationException(
                    "Failed to unprotect document: The password may be incorrect. Please verify the password and try again.");

            doc.Save(outputPath);
            return $"Protection removed successfully (was: {previousProtectionType})\nOutput: {outputPath}";
        });
    }
}