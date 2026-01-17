using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint page setup (slide size, orientation, footer, slide numbering).
///     Merges: PptSetSlideSizeTool, PptSetSlideOrientationTool, PptHeaderFooterTool
/// </summary>
[McpServerToolType]
public class PptPageSetupTool
{
    /// <summary>
    ///     Handler registry for page setup operations.
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
    ///     Initializes a new instance of the <see cref="PptPageSetupTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptPageSetupTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.PageSetup");
    }

    /// <summary>
    ///     Executes a PowerPoint page setup operation (set_size, set_orientation, set_footer, set_slide_numbering).
    /// </summary>
    /// <param name="operation">The operation to perform: set_size, set_orientation, set_footer, set_slide_numbering.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="preset">Preset: OnScreen16x9, OnScreen16x10, Letter, A4, Banner, Custom (optional, for set_size).</param>
    /// <param name="width">Custom width in points when preset=Custom (1-5000, 1 inch = 72 points).</param>
    /// <param name="height">Custom height in points when preset=Custom (1-5000, 1 inch = 72 points).</param>
    /// <param name="orientation">Orientation: Portrait or Landscape (required for set_orientation).</param>
    /// <param name="footerText">Footer text (optional, for set_footer).</param>
    /// <param name="dateText">Date/time text (optional, for set_footer).</param>
    /// <param name="showSlideNumber">Show slide number (optional, for set_footer/set_slide_numbering, default: true).</param>
    /// <param name="firstNumber">First slide number (optional, for set_slide_numbering, default: 1).</param>
    /// <param name="slideIndices">Slide indices (0-based, optional, for set_footer, if not provided applies to all slides).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_page_setup")]
    [Description(
        @"Manage PowerPoint page setup. Supports 4 operations: set_size, set_orientation, set_footer, set_slide_numbering.

Note: PowerPoint slides do not have a separate header field. Only footer, date, and slide number are available.
Size unit: 1 inch = 72 points. Valid range: 1-5000 points.

Usage examples:
- Set slide size: ppt_page_setup(operation='set_size', path='presentation.pptx', preset='OnScreen16x9')
- Set custom size: ppt_page_setup(operation='set_size', path='presentation.pptx', preset='Custom', width=960, height=720)
- Set orientation: ppt_page_setup(operation='set_orientation', path='presentation.pptx', orientation='Portrait')
- Set footer: ppt_page_setup(operation='set_footer', path='presentation.pptx', footerText='Footer', showSlideNumber=true)
- Set slide numbering: ppt_page_setup(operation='set_slide_numbering', path='presentation.pptx', showSlideNumber=true, firstNumber=1)")]
    public string Execute(
        [Description("Operation: set_size, set_orientation, set_footer, set_slide_numbering")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Preset: OnScreen16x9, OnScreen16x10, Letter, A4, Banner, Custom (optional, for set_size)")]
        string? preset = null,
        [Description("Custom width in points when preset=Custom (1-5000, 1 inch = 72 points)")]
        double? width = null,
        [Description("Custom height in points when preset=Custom (1-5000, 1 inch = 72 points)")]
        double? height = null,
        [Description("Orientation: 'Portrait' or 'Landscape' (required for set_orientation)")]
        string? orientation = null,
        [Description("Footer text (optional, for set_footer)")]
        string? footerText = null,
        [Description("Date/time text (optional, for set_footer)")]
        string? dateText = null,
        [Description("Show slide number (optional, for set_footer/set_slide_numbering, default: true)")]
        bool showSlideNumber = true,
        [Description("First slide number (optional, for set_slide_numbering, default: 1)")]
        int firstNumber = 1,
        [Description("Slide indices (0-based, optional, for set_footer, if not provided applies to all slides)")]
        int[]? slideIndices = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, preset, width, height, orientation,
            footerText, dateText, showSlideNumber, firstNumber, slideIndices);

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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? preset,
        double? width,
        double? height,
        string? orientation,
        string? footerText,
        string? dateText,
        bool showSlideNumber,
        int firstNumber,
        int[]? slideIndices)
    {
        return operation.ToLowerInvariant() switch
        {
            "set_size" => BuildSetSizeParameters(preset, width, height),
            "set_orientation" => BuildSetOrientationParameters(orientation),
            "set_footer" => BuildSetFooterParameters(footerText, dateText, showSlideNumber, slideIndices),
            "set_slide_numbering" => BuildSetSlideNumberingParameters(showSlideNumber, firstNumber),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the set slide size operation.
    /// </summary>
    /// <param name="preset">The preset size (OnScreen16x9, OnScreen16x10, Letter, A4, Banner, Custom).</param>
    /// <param name="width">The custom width in points when preset is Custom.</param>
    /// <param name="height">The custom height in points when preset is Custom.</param>
    /// <returns>OperationParameters configured for setting slide size.</returns>
    private static OperationParameters BuildSetSizeParameters(string? preset, double? width, double? height)
    {
        var parameters = new OperationParameters();
        if (preset != null) parameters.Set("preset", preset);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set slide orientation operation.
    /// </summary>
    /// <param name="orientation">The orientation (Portrait or Landscape).</param>
    /// <returns>OperationParameters configured for setting slide orientation.</returns>
    private static OperationParameters BuildSetOrientationParameters(string? orientation)
    {
        var parameters = new OperationParameters();
        if (orientation != null) parameters.Set("orientation", orientation);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set footer operation.
    /// </summary>
    /// <param name="footerText">The footer text.</param>
    /// <param name="dateText">The date/time text.</param>
    /// <param name="showSlideNumber">Whether to show slide number.</param>
    /// <param name="slideIndices">The slide indices to apply footer to.</param>
    /// <returns>OperationParameters configured for setting footer.</returns>
    private static OperationParameters BuildSetFooterParameters(string? footerText, string? dateText,
        bool showSlideNumber, int[]? slideIndices)
    {
        var parameters = new OperationParameters();
        if (footerText != null) parameters.Set("footerText", footerText);
        if (dateText != null) parameters.Set("dateText", dateText);
        parameters.Set("showSlideNumber", showSlideNumber);
        if (slideIndices != null) parameters.Set("slideIndices", slideIndices);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set slide numbering operation.
    /// </summary>
    /// <param name="showSlideNumber">Whether to show slide number.</param>
    /// <param name="firstNumber">The first slide number.</param>
    /// <returns>OperationParameters configured for setting slide numbering.</returns>
    private static OperationParameters BuildSetSlideNumberingParameters(bool showSlideNumber, int firstNumber)
    {
        var parameters = new OperationParameters();
        parameters.Set("showSlideNumber", showSlideNumber);
        parameters.Set("firstNumber", firstNumber);
        return parameters;
    }
}
