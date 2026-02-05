using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint fonts (replace, embed, remove_embedded, get_used, set_fallback).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Font")]
[McpServerToolType]
public class PptFontTool
{
    /// <summary>
    ///     Handler registry for font operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptFontTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptFontTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Font");
    }

    /// <summary>
    ///     Executes a PowerPoint font operation (replace, embed, remove_embedded, get_used, set_fallback).
    /// </summary>
    /// <param name="operation">The operation to perform: replace, embed, remove_embedded, get_used, set_fallback.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sourceFont">Source font name (for replace).</param>
    /// <param name="targetFont">Target font name (for replace).</param>
    /// <param name="fontName">Font name (for embed, remove_embedded).</param>
    /// <param name="fallbackFont">Fallback font name (for set_fallback).</param>
    /// <param name="embedMode">Embed mode: 'all' or 'subset' (for embed, default: all).</param>
    /// <param name="unicodeStart">Unicode range start (for set_fallback, default: 0x0000).</param>
    /// <param name="unicodeEnd">Unicode range end (for set_fallback, default: 0xFFFF).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_font",
        Title = "PowerPoint Font Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage PowerPoint fonts. Supports 5 operations: replace, embed, remove_embedded, get_used, set_fallback.

Usage examples:
- Replace font: ppt_font(operation='replace', path='file.pptx', sourceFont='Arial', targetFont='Calibri')
- Embed font: ppt_font(operation='embed', path='file.pptx', fontName='Custom Font')
- Remove embedded: ppt_font(operation='remove_embedded', path='file.pptx', fontName='Custom Font')
- Get used fonts: ppt_font(operation='get_used', path='file.pptx')
- Set fallback: ppt_font(operation='set_fallback', path='file.pptx', fallbackFont='Arial')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'replace': Replace font (required: sourceFont, targetFont)
- 'embed': Embed font (required: fontName)
- 'remove_embedded': Remove embedded font (required: fontName)
- 'get_used': List all used fonts
- 'set_fallback': Set font fallback rule (required: fallbackFont)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional)")]
        string? outputPath = null,
        [Description("Source font name (for replace)")]
        string? sourceFont = null,
        [Description("Target font name (for replace)")]
        string? targetFont = null,
        [Description("Font name (for embed, remove_embedded)")]
        string? fontName = null,
        [Description("Fallback font name (for set_fallback)")]
        string? fallbackFont = null,
        [Description("Embed mode: 'all' or 'subset' (for embed, default: all)")]
        string embedMode = "all",
        [Description("Unicode range start (for set_fallback, default: 0x0000)")]
        int unicodeStart = 0x0000,
        [Description("Unicode range end (for set_fallback, default: 0xFFFF)")]
        int unicodeEnd = 0xFFFF)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters =
            BuildParameters(operation, sourceFont, targetFont, fontName, fallbackFont, embedMode, unicodeStart,
                unicodeEnd);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get_used", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(string operation, string? sourceFont, string? targetFont,
        string? fontName, string? fallbackFont, string embedMode, int unicodeStart, int unicodeEnd)
    {
        return operation.ToLowerInvariant() switch
        {
            "replace" => BuildReplaceParameters(sourceFont, targetFont),
            "embed" => BuildEmbedParameters(fontName, embedMode),
            "remove_embedded" => BuildRemoveParameters(fontName),
            "set_fallback" => BuildFallbackParameters(fallbackFont, unicodeStart, unicodeEnd),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the replace font operation.
    /// </summary>
    /// <param name="sourceFont">The source font name.</param>
    /// <param name="targetFont">The target font name.</param>
    /// <returns>OperationParameters configured for replacing a font.</returns>
    private static OperationParameters BuildReplaceParameters(string? sourceFont, string? targetFont)
    {
        var parameters = new OperationParameters();
        if (sourceFont != null) parameters.Set("sourceFont", sourceFont);
        if (targetFont != null) parameters.Set("targetFont", targetFont);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the embed font operation.
    /// </summary>
    /// <param name="fontName">The font name to embed.</param>
    /// <param name="embedMode">The embed mode ('all' or 'subset').</param>
    /// <returns>OperationParameters configured for embedding a font.</returns>
    private static OperationParameters BuildEmbedParameters(string? fontName, string embedMode)
    {
        var parameters = new OperationParameters();
        if (fontName != null) parameters.Set("fontName", fontName);
        parameters.Set("embedMode", embedMode);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the remove embedded font operation.
    /// </summary>
    /// <param name="fontName">The font name to remove.</param>
    /// <returns>OperationParameters configured for removing an embedded font.</returns>
    private static OperationParameters BuildRemoveParameters(string? fontName)
    {
        var parameters = new OperationParameters();
        if (fontName != null) parameters.Set("fontName", fontName);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set fallback font operation.
    /// </summary>
    /// <param name="fallbackFont">The fallback font name.</param>
    /// <param name="unicodeStart">The Unicode range start.</param>
    /// <param name="unicodeEnd">The Unicode range end.</param>
    /// <returns>OperationParameters configured for setting a font fallback rule.</returns>
    private static OperationParameters BuildFallbackParameters(string? fallbackFont, int unicodeStart, int unicodeEnd)
    {
        var parameters = new OperationParameters();
        if (fallbackFont != null) parameters.Set("fallbackFont", fallbackFont);
        parameters.Set("unicodeStart", unicodeStart);
        parameters.Set("unicodeEnd", unicodeEnd);
        return parameters;
    }
}
