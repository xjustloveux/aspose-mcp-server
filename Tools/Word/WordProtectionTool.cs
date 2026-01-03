using System.ComponentModel;
using Aspose.Words;
using Aspose.Words.Loading;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document protection (protect, unprotect)
///     Merges: WordProtectTool, WordUnprotectTool
/// </summary>
[McpServerToolType]
public class WordProtectionTool
{
    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordProtectionTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordProtectionTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_protection")]
    [Description(@"Protect or unprotect a Word document. Supports 2 operations: protect, unprotect.

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
- For encrypted documents (with open password), the same password will be used to open the file")]
    public string Execute(
        [Description("Operation: protect, unprotect")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (if not provided, overwrites input)")]
        string? outputPath = null,
        [Description("Protection password (required for protect operation, optional for unprotect)")]
        string? password = null,
        [Description(
            "Protection type: 'ReadOnly', 'AllowOnlyComments', 'AllowOnlyFormFields', 'AllowOnlyRevisions' (required for protect operation)")]
        string protectionType = "ReadOnly")
    {
        // For protection operations, we need special handling because encrypted files need password to open
        // So we don't use DocumentContext for protect/unprotect operations directly
        return operation.ToLower() switch
        {
            "protect" => Protect(path, sessionId, outputPath, password, protectionType),
            "unprotect" => Unprotect(path, sessionId, outputPath, password),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Protects the document with specified protection type and password.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="password">The protection password.</param>
    /// <param name="protectionType">The protection type string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when password is null or empty, or path/sessionId are not provided.</exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    private string Protect(string? path, string? sessionId, string? outputPath, string? password, string protectionType)
    {
        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException(
                "Password is required for protect operation. Please provide a non-empty password.");

        if (!string.IsNullOrEmpty(sessionId))
        {
            // Session mode
            if (_sessionManager == null)
                throw new InvalidOperationException(
                    "Session management is not enabled. Use --enable-sessions flag or provide a file path.");

            var doc = _sessionManager.GetDocument<Document>(sessionId);
            var protectionTypeEnum = GetProtectionType(protectionType);
            doc.Protect(protectionTypeEnum, password);
            _sessionManager.MarkDirty(sessionId);

            return
                $"Document protected with {protectionTypeEnum}. Changes applied to session {sessionId}. Use document_session(operation='save') to save to disk.";
        }

        // File mode
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Either sessionId or path must be provided");

        var document = LoadDocument(path, password);
        var protType = GetProtectionType(protectionType);

        document.Protect(protType, password);
        var savePath = outputPath ?? path;
        document.Save(savePath);

        return $"Document protected with {protType}: {savePath}";
    }

    /// <summary>
    ///     Loads a Word document with optional password for encrypted files.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="password">The optional password for encrypted documents.</param>
    /// <returns>The loaded Word document.</returns>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the document cannot be loaded due to file not found, incorrect
    ///     password, unsupported format, or other errors.
    /// </exception>
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
    ///     Converts a protection type string to ProtectionType enum using Enum.TryParse.
    /// </summary>
    /// <param name="protectionTypeStr">The protection type string to parse.</param>
    /// <returns>The parsed ProtectionType enum value, or ReadOnly if parsing fails.</returns>
    private static ProtectionType GetProtectionType(string protectionTypeStr)
    {
        if (Enum.TryParse<ProtectionType>(protectionTypeStr, true, out var result))
            return result;

        return ProtectionType.ReadOnly;
    }

    /// <summary>
    ///     Removes protection from the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="password">The optional password for the protected document.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither path nor sessionId is provided.</exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled or unprotection fails.</exception>
    private string Unprotect(string? path, string? sessionId, string? outputPath, string? password)
    {
        if (!string.IsNullOrEmpty(sessionId))
        {
            // Session mode
            if (_sessionManager == null)
                throw new InvalidOperationException(
                    "Session management is not enabled. Use --enable-sessions flag or provide a file path.");

            var doc = _sessionManager.GetDocument<Document>(sessionId);
            var previousProtectionType = doc.ProtectionType;

            if (previousProtectionType == ProtectionType.NoProtection)
                return "Document is not protected, no need to unprotect";

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

            _sessionManager.MarkDirty(sessionId);
            return
                $"Protection removed successfully (was: {previousProtectionType}). Changes applied to session {sessionId}. Use document_session(operation='save') to save to disk.";
        }

        // File mode
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Either sessionId or path must be provided");

        var document = LoadDocument(path, password);
        var prevProtectionType = document.ProtectionType;

        if (prevProtectionType == ProtectionType.NoProtection)
        {
            var savePath = outputPath ?? path;
            if (!string.Equals(path, savePath, StringComparison.OrdinalIgnoreCase))
            {
                document.Save(savePath);
                return $"Document is not protected, saved to: {savePath}";
            }

            return "Document is not protected, no need to unprotect";
        }

        try
        {
            document.Unprotect(password);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to unprotect document: The password may be incorrect or the document has additional restrictions. Details: {ex.Message}",
                ex);
        }

        if (document.ProtectionType != ProtectionType.NoProtection)
            throw new InvalidOperationException(
                "Failed to unprotect document: The password may be incorrect. Please verify the password and try again.");

        var finalOutputPath = outputPath ?? path;
        document.Save(finalOutputPath);
        return $"Protection removed successfully (was: {prevProtectionType})\nOutput: {finalOutputPath}";
    }
}