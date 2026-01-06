using System.ComponentModel;
using Aspose.Words;
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
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordProtectionTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordProtectionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word protection operation (protect, unprotect).
    /// </summary>
    /// <param name="operation">The operation to perform: protect, unprotect.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (if not provided, overwrites input).</param>
    /// <param name="password">Protection password (required for protect, optional for unprotect).</param>
    /// <param name="protectionType">Protection type: ReadOnly, AllowOnlyComments, AllowOnlyFormFields, AllowOnlyRevisions.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when password is missing for protect operation or the operation is unknown.</exception>
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
    /// <exception cref="ArgumentException">Thrown when password is null or empty.</exception>
    private string Protect(string? path, string? sessionId, string? outputPath, string? password, string protectionType)
    {
        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException(
                "Password is required for protect operation. Please provide a non-empty password.");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor, password);
        var doc = ctx.Document;
        var protectionTypeEnum = GetProtectionType(protectionType);

        doc.Protect(protectionTypeEnum, password);
        ctx.Save(outputPath);

        return ctx.IsSession
            ? $"Document protected with {protectionTypeEnum}. {ctx.GetOutputMessage()}"
            : $"Document protected with {protectionTypeEnum}: {outputPath ?? path}";
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
    /// <exception cref="InvalidOperationException">Thrown when unprotection fails.</exception>
    private string Unprotect(string? path, string? sessionId, string? outputPath, string? password)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor, password);
        var doc = ctx.Document;
        var previousProtectionType = doc.ProtectionType;

        if (previousProtectionType == ProtectionType.NoProtection)
        {
            if (!ctx.IsSession && outputPath != null &&
                !string.Equals(path, outputPath, StringComparison.OrdinalIgnoreCase))
            {
                ctx.Save(outputPath);
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

        ctx.Save(outputPath);

        return ctx.IsSession
            ? $"Protection removed successfully (was: {previousProtectionType}). {ctx.GetOutputMessage()}"
            : $"Protection removed successfully (was: {previousProtectionType})\nOutput: {outputPath ?? path}";
    }
}