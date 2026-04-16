using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Ole;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     MCP tool exposing Excel-workbook OLE-object operations (<c>list</c> /
///     <c>extract</c> / <c>extract_all</c> / <c>remove</c>) over Aspose.Cells'
///     <see cref="Aspose.Cells.Drawing.OleObject" /> surface. Accepts <c>.xlsx</c> and
///     <c>.xls</c>.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.OleObject")]
[McpServerToolType]
public class ExcelOleObjectTool
{
    /// <summary>Registry for the four Excel OLE handlers.</summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>Optional session identity accessor.</summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>Optional server configuration.</summary>
    private readonly ServerConfig? _serverConfig;

    /// <summary>Optional unified session manager.</summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>Initializes a new instance of the <see cref="ExcelOleObjectTool" /> class.</summary>
    /// <param name="sessionManager">Optional session manager.</param>
    /// <param name="identityAccessor">Optional identity accessor.</param>
    /// <param name="serverConfig">Optional server config (allowlist, byte cap).</param>
    public ExcelOleObjectTool(
        DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null,
        ServerConfig? serverConfig = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _serverConfig = serverConfig;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace(
            "AsposeMcpServer.Handlers.Excel.OleObject");
    }

    /// <summary>
    ///     Executes an Excel OLE-object operation (see
    ///     <see cref="AsposeMcpServer.Tools.Word.WordOleObjectTool.Execute" /> for shape).
    /// </summary>
    /// <param name="operation">One of <c>list</c> / <c>extract</c> / <c>extract_all</c> / <c>remove</c>.</param>
    /// <param name="path">Source file path.</param>
    /// <param name="sessionId">Session ID.</param>
    /// <param name="password">File-mode password.</param>
    /// <param name="outputDirectory">Destination directory.</param>
    /// <param name="oleIndex">Zero-based OLE index (flat across sheets).</param>
    /// <param name="outputFileName">Filename override for <c>extract</c>.</param>
    /// <param name="outputPath">Re-save target for file-mode <c>remove</c>.</param>
    /// <returns>Finalized operation result.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the operation is not recognized or the source path fails shape
    ///     validation.
    /// </exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when the source path falls outside the configured allowlist.</exception>
    [McpServerTool(
        Name = "excel_ole_object",
        Title = "Excel OLE Object Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        "Manage OLE objects embedded in Excel workbooks (.xlsx, .xls). Supports 4 operations: list, extract, extract_all, remove.\n\n" +
        "Indices are flat across all worksheets (sheet 0 OLE 0 = global 0, sheet 1 OLE 0 = global N). Remove semantics: removing index N shifts indices > N down by one (AC-17). " +
        "File-mode remove is best-effort last-writer-wins — use session-mode for concurrent-safe remove.")]
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

            OleExtensionGuard.EnsureExcelExtension(path);
        }

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor, password);
        var passwordIgnored = !string.IsNullOrEmpty(sessionId) && !string.IsNullOrEmpty(password);

        var parameters = BuildParameters(outputDirectory, oleIndex, outputFileName, outputPath);
        var opCtx = new OperationContext<Workbook>
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

    /// <summary>Packs the optional parameters into an <see cref="OperationParameters" />.</summary>
    /// <param name="outputDirectory">Destination directory.</param>
    /// <param name="oleIndex">Zero-based OLE index.</param>
    /// <param name="outputFileName">Filename override.</param>
    /// <param name="outputPath">Re-save target.</param>
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
