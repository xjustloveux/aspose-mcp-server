using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Ole;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     MCP tool exposing Word-document OLE-object operations (<c>list</c> / <c>extract</c> /
///     <c>extract_all</c> / <c>remove</c>) over Aspose.Words'
///     <see cref="Aspose.Words.Drawing.OleFormat" /> /
///     <see cref="Aspose.Words.Drawing.OlePackage" /> surface. Accepts <c>.docx</c> and
///     <c>.doc</c>; legacy binary storage is handled uniformly by Aspose's object model.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.OleObject")]
[McpServerToolType]
public class WordOleObjectTool
{
    /// <summary>Registry for the four Word OLE handlers.</summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>Optional session identity accessor for session isolation.</summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>Optional server configuration (allowlist, cumulative-byte cap).</summary>
    private readonly ServerConfig? _serverConfig;

    /// <summary>Optional unified session manager.</summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WordOleObjectTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for session-mode access.</param>
    /// <param name="identityAccessor">Optional session identity accessor.</param>
    /// <param name="serverConfig">
    ///     Optional server configuration — supplies the allowed-base-paths list and the
    ///     cumulative-byte cap (<see cref="ServerConfig.MaxExtractAllBytes" />).
    /// </param>
    public WordOleObjectTool(
        DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null,
        ServerConfig? serverConfig = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _serverConfig = serverConfig;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace(
            "AsposeMcpServer.Handlers.Word.OleObject");
    }

    /// <summary>
    ///     Executes a Word OLE-object operation.
    /// </summary>
    /// <param name="operation">One of <c>list</c> / <c>extract</c> / <c>extract_all</c> / <c>remove</c>.</param>
    /// <param name="path">Source file path (required when no <paramref name="sessionId" />).</param>
    /// <param name="sessionId">Session ID (alternative to <paramref name="path" />).</param>
    /// <param name="password">File-mode password for protected documents (ignored in session-mode with a note).</param>
    /// <param name="outputDirectory">Destination directory for <c>extract</c> / <c>extract_all</c>.</param>
    /// <param name="oleIndex">Zero-based OLE index (required for <c>extract</c> and <c>remove</c>).</param>
    /// <param name="outputFileName">Optional sanitized filename override for <c>extract</c>.</param>
    /// <param name="outputPath">Optional re-save target for file-mode <c>remove</c>; defaults to <paramref name="path" />.</param>
    /// <returns>Operation-specific result wrapped by <see cref="ResultHelper.FinalizeResult{TDoc, TResult}" />.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the operation is not recognized or the source path fails shape
    ///     validation.
    /// </exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when the source path falls outside the configured allowlist.</exception>
    [McpServerTool(
        Name = "word_ole_object",
        Title = "Word OLE Object Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        "Manage OLE objects embedded in Word documents (.docx, .doc). Supports 4 operations: list, extract, extract_all, remove.\n\n" +
        "Usage examples:\n" +
        "- List OLE objects: word_ole_object(operation='list', path='doc.docx')\n" +
        "- Extract by index: word_ole_object(operation='extract', path='doc.docx', outputDirectory='./out', oleIndex=0)\n" +
        "- Extract all: word_ole_object(operation='extract_all', path='doc.docx', outputDirectory='./out')\n" +
        "- Remove by index: word_ole_object(operation='remove', path='doc.docx', oleIndex=0)\n\n" +
        "Remove semantics: removing index N shifts indices > N down by one (AC-17). File-mode remove is best-effort last-writer-wins: " +
        "two concurrent removes on the same path may each load a pre-remove snapshot and each save, silently erasing the earlier change. " +
        "Use session-mode for concurrent-safe remove semantics.")]
    public object Execute(
        [Description("Operation: list | extract | extract_all | remove")]
        string operation,
        [Description("Source file path; required when sessionId is null")]
        string? path = null,
        [Description("Session ID; alternative to path for in-memory editing")]
        string? sessionId = null,
        [Description("Password for protected source files (file-mode only; ignored in session-mode)")]
        string? password = null,
        [Description("Output directory for extract / extract_all")]
        string? outputDirectory = null,
        [Description("Zero-based OLE index (required for extract + remove)")]
        int? oleIndex = null,
        [Description("Optional sanitized filename override for extract")]
        string? outputFileName = null,
        [Description("Optional output path for re-saving after remove in file-mode")]
        string? outputPath = null)
    {
        if (!_handlerRegistry.HasHandler(operation))
            throw new ArgumentException(OleErrorMessageBuilder.UnknownOperation(operation), nameof(operation));

        if (string.IsNullOrEmpty(sessionId) && !string.IsNullOrEmpty(path))
        {
            try
            {
                SecurityHelper.ValidateFilePath(path, nameof(path), true);
            }
            catch (ArgumentException)
            {
                throw new ArgumentException(OleErrorMessageBuilder.InvalidPath(path), nameof(path));
            }

            if (_serverConfig is { AllowedBasePaths.Count: > 0 })
                try
                {
                    SecurityHelper.ValidatePathWithinAllowedBases(path, _serverConfig.AllowedBasePaths);
                }
                catch (ArgumentException)
                {
                    throw new UnauthorizedAccessException(OleErrorMessageBuilder.InvalidPath(path));
                }

            OleExtensionGuard.EnsureWordExtension(path);
        }

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor, password);
        var passwordIgnored = !string.IsNullOrEmpty(sessionId) && !string.IsNullOrEmpty(password);

        var parameters = BuildParameters(outputDirectory, oleIndex, outputFileName, outputPath);
        var opCtx = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath,
            ServerConfig = _serverConfig
        };

        var handler = _handlerRegistry.GetHandler(operation);
        var result = handler.Execute(opCtx, parameters);
        result = OleToolHelper.AttachPasswordIgnoredNote(result, passwordIgnored);

        if (opCtx.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Packs the optional extract / extract_all / remove parameters into an
    ///     <see cref="OperationParameters" /> bag using the shared key constants.
    /// </summary>
    /// <param name="outputDirectory">Destination directory (nullable).</param>
    /// <param name="oleIndex">Zero-based OLE index (nullable).</param>
    /// <param name="outputFileName">Filename override (nullable).</param>
    /// <param name="outputPath">Re-save target (nullable).</param>
    /// <returns>A fully-populated <see cref="OperationParameters" />.</returns>
    private static OperationParameters BuildParameters(
        string? outputDirectory, int? oleIndex, string? outputFileName, string? outputPath)
    {
        var parameters = new OperationParameters();
        parameters.SetIfNotNull(OleParamKeys.OutputDirectory, outputDirectory);
        parameters.SetIfHasValue(OleParamKeys.OleIndex, oleIndex);
        parameters.SetIfNotNull(OleParamKeys.OutputFileName, outputFileName);
        parameters.SetIfNotNull(OleParamKeys.OutputPath, outputPath);
        return parameters;
    }
}
