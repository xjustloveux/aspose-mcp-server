using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
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
    ///     Minimum allowed slide size in points.
    /// </summary>
    private const float MinSizePoints = 1f;

    /// <summary>
    ///     Maximum allowed slide size in points.
    /// </summary>
    private const float MaxSizePoints = 5000f;

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

        return operation.ToLower() switch
        {
            "set_size" => SetSlideSize(ctx, outputPath, preset, width, height),
            "set_orientation" => SetSlideOrientation(ctx, outputPath, orientation),
            "set_footer" => SetFooter(ctx, outputPath, footerText, showSlideNumber, dateText, slideIndices),
            "set_slide_numbering" => SetSlideNumbering(ctx, outputPath, showSlideNumber, firstNumber),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets slide size.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="preset">The slide size preset name.</param>
    /// <param name="width">The custom width in points.</param>
    /// <param name="height">The custom height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when custom size is selected but width or height is not provided.</exception>
    private static string SetSlideSize(DocumentContext<Presentation> ctx, string? outputPath, string? preset,
        double? width, double? height)
    {
        var presetValue = preset ?? "OnScreen16x9";
        var presentation = ctx.Document;
        var slideSize = presentation.SlideSize;
        var type = presetValue.ToLower() switch
        {
            "onscreen16x10" => SlideSizeType.OnScreen16x10,
            "a4" => SlideSizeType.A4Paper,
            "banner" => SlideSizeType.Banner,
            "custom" => SlideSizeType.Custom,
            _ => SlideSizeType.OnScreen
        };

        if (type == SlideSizeType.Custom)
        {
            if (!width.HasValue || !height.HasValue)
                throw new ArgumentException("Custom size requires width and height.");

            ValidateSizeRange(width.Value, height.Value);
            slideSize.SetSize((float)width.Value, (float)height.Value, SlideSizeScaleType.DoNotScale);
        }
        else
        {
            slideSize.SetSize(type, SlideSizeScaleType.DoNotScale);
        }

        ctx.Save(outputPath);
        var sizeInfo = slideSize.Type == SlideSizeType.Custom
            ? $" ({slideSize.Size.Width}x{slideSize.Size.Height})"
            : "";

        var result = $"Slide size set to {slideSize.Type}{sizeInfo}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Validates that width and height are within acceptable range.
    /// </summary>
    /// <param name="width">The width value to validate.</param>
    /// <param name="height">The height value to validate.</param>
    /// <exception cref="ArgumentException">Thrown when width or height is outside the valid range.</exception>
    private static void ValidateSizeRange(double width, double height)
    {
        if (width < MinSizePoints || width > MaxSizePoints)
            throw new ArgumentException($"Width must be between {MinSizePoints} and {MaxSizePoints} points.");

        if (height < MinSizePoints || height > MaxSizePoints)
            throw new ArgumentException($"Height must be between {MinSizePoints} and {MaxSizePoints} points.");
    }

    /// <summary>
    ///     Sets slide orientation by swapping width and height while preserving the aspect ratio.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="orientation">The orientation value (Portrait or Landscape).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when orientation is not provided.</exception>
    private static string SetSlideOrientation(DocumentContext<Presentation> ctx, string? outputPath,
        string? orientation)
    {
        if (string.IsNullOrEmpty(orientation))
            throw new ArgumentException("orientation is required for set_orientation operation");

        var isPortrait = orientation.Equals("Portrait", StringComparison.OrdinalIgnoreCase);
        var presentation = ctx.Document;
        var currentSize = presentation.SlideSize.Size;
        var currentWidth = currentSize.Width;
        var currentHeight = currentSize.Height;

        var needsSwap = isPortrait ? currentWidth > currentHeight : currentHeight > currentWidth;

        if (needsSwap)
            presentation.SlideSize.SetSize(currentHeight, currentWidth, SlideSizeScaleType.EnsureFit);

        ctx.Save(outputPath);

        var finalSize = presentation.SlideSize.Size;
        var result = $"Slide orientation set to {orientation} ({finalSize.Width}x{finalSize.Height}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets footer text, date, and slide number for slides.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="footerText">The footer text.</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="dateText">The date/time text.</param>
    /// <param name="slideIndices">The slide indices to apply footer to.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetFooter(DocumentContext<Presentation> ctx, string? outputPath,
        string? footerText, bool showSlideNumber, string? dateText, int[]? slideIndices)
    {
        var presentation = ctx.Document;
        var slides = GetTargetSlides(presentation, slideIndices);
        var applyToAll = slideIndices == null || slideIndices.Length == 0;

        if (applyToAll)
            EnableMasterVisibility(presentation, footerText, showSlideNumber, dateText);

        foreach (var slide in slides)
            ApplyFooterSettings(slide.HeaderFooterManager, footerText, showSlideNumber, dateText);

        ctx.Save(outputPath);

        var result = $"Footer settings updated for {slides.Count} slide(s).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets slide numbering visibility and start number.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="firstNumber">The first slide number.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetSlideNumbering(DocumentContext<Presentation> ctx, string? outputPath, bool showSlideNumber,
        int firstNumber)
    {
        var presentation = ctx.Document;

        presentation.FirstSlideNumber = firstNumber;
        presentation.HeaderFooterManager.SetAllSlideNumbersVisibility(showSlideNumber);

        foreach (var slide in presentation.Slides)
            slide.HeaderFooterManager.SetSlideNumberVisibility(showSlideNumber);

        ctx.Save(outputPath);

        var visibilityText = showSlideNumber ? "shown" : "hidden";
        var result = $"Slide numbers {visibilityText}, starting from {firstNumber}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets target slides based on slide indices.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="slideIndices">The slide indices to retrieve.</param>
    /// <returns>A list of slides matching the specified indices, or all slides if indices is null or empty.</returns>
    private static List<ISlide> GetTargetSlides(IPresentation presentation, int[]? slideIndices)
    {
        if (slideIndices == null || slideIndices.Length == 0)
            return presentation.Slides.ToList();

        List<ISlide> slides = [];
        foreach (var index in slideIndices)
        {
            PowerPointHelper.ValidateSlideIndex(index, presentation);
            slides.Add(presentation.Slides[index]);
        }

        return slides;
    }

    /// <summary>
    ///     Enables visibility on master slides.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="footerText">The footer text.</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="dateText">The date/time text.</param>
    private static void EnableMasterVisibility(IPresentation presentation, string? footerText, bool showSlideNumber,
        string? dateText)
    {
        var manager = presentation.HeaderFooterManager;

        if (!string.IsNullOrEmpty(footerText))
            manager.SetAllFootersVisibility(true);

        manager.SetAllSlideNumbersVisibility(showSlideNumber);

        if (!string.IsNullOrEmpty(dateText))
            manager.SetAllDateTimesVisibility(true);
    }

    /// <summary>
    ///     Applies footer settings to a slide.
    /// </summary>
    /// <param name="manager">The slide header footer manager.</param>
    /// <param name="footerText">The footer text.</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="dateText">The date/time text.</param>
    private static void ApplyFooterSettings(ISlideHeaderFooterManager manager, string? footerText,
        bool showSlideNumber, string? dateText)
    {
        if (!string.IsNullOrEmpty(footerText))
        {
            manager.SetFooterText(footerText);
            manager.SetFooterVisibility(true);
        }
        else
        {
            manager.SetFooterVisibility(false);
        }

        manager.SetSlideNumberVisibility(showSlideNumber);

        if (!string.IsNullOrEmpty(dateText))
        {
            manager.SetDateTimeText(dateText);
            manager.SetDateTimeVisibility(true);
        }
        else
        {
            manager.SetDateTimeVisibility(false);
        }
    }
}