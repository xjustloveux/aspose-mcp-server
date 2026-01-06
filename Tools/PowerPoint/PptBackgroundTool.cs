using System.ComponentModel;
using System.Drawing;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint backgrounds (set, get).
/// </summary>
[McpServerToolType]
public class PptBackgroundTool
{
    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptBackgroundTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptBackgroundTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint background operation (set, get).
    /// </summary>
    /// <param name="operation">The operation to perform: set, get.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based, default: 0, ignored if applyToAll is true).</param>
    /// <param name="color">Hex color like #FFAA00 or #80FFAA00 (with alpha).</param>
    /// <param name="imagePath">Background image path.</param>
    /// <param name="applyToAll">Apply background to all slides (default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_background")]
    [Description(@"Manage PowerPoint backgrounds. Supports 2 operations: set, get.

Usage examples:
- Set background color: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, color='#FFFFFF')
- Set background image: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, imagePath='bg.png')
- Apply to all slides: ppt_background(operation='set', path='presentation.pptx', color='#FFFFFF', applyToAll=true)
- Get background: ppt_background(operation='get', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description("Operation: set, get")] string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, default: 0, ignored if applyToAll is true)")]
        int slideIndex = 0,
        [Description("Hex color like #FFAA00 or #80FFAA00 (with alpha)")]
        string? color = null,
        [Description("Background image path")] string? imagePath = null,
        [Description("Apply background to all slides (default: false)")]
        bool applyToAll = false)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "set" => SetBackground(ctx, outputPath, slideIndex, color, imagePath, applyToAll),
            "get" => GetBackground(ctx, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets slide background with color or image.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="colorHex">The hex color string.</param>
    /// <param name="imagePath">The background image path.</param>
    /// <param name="applyToAll">Whether to apply to all slides.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither color nor imagePath is provided.</exception>
    private static string SetBackground(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        string? colorHex, string? imagePath, bool applyToAll)
    {
        if (string.IsNullOrWhiteSpace(colorHex) && string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Please provide at least one of color or imagePath");

        var presentation = ctx.Document;

        IPPImage? img = null;
        if (!string.IsNullOrWhiteSpace(imagePath))
            img = presentation.Images.AddImage(File.ReadAllBytes(imagePath));

        Color? color = null;
        if (!string.IsNullOrWhiteSpace(colorHex))
            color = ColorHelper.ParseColor(colorHex);

        var slidesToUpdate = applyToAll
            ? presentation.Slides.ToList()
            : [PowerPointHelper.GetSlide(presentation, slideIndex)];

        foreach (var slide in slidesToUpdate)
            ApplyBackground(slide, color, img);

        ctx.Save(outputPath);

        var message = applyToAll
            ? $"Background applied to all {slidesToUpdate.Count} slides"
            : $"Background updated for slide {slideIndex}";
        var result = $"{message}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Applies background color or image to a slide.
    /// </summary>
    /// <param name="slide">The slide to apply background to.</param>
    /// <param name="color">The background color.</param>
    /// <param name="image">The background image.</param>
    private static void ApplyBackground(ISlide slide, Color? color, IPPImage? image)
    {
        slide.Background.Type = BackgroundType.OwnBackground;
        var fillFormat = slide.Background.FillFormat;

        if (image != null)
        {
            fillFormat.FillType = FillType.Picture;
            fillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
            fillFormat.PictureFillFormat.Picture.Image = image;
        }
        else if (color.HasValue)
        {
            fillFormat.FillType = FillType.Solid;
            fillFormat.SolidFillColor.Color = color.Value;
        }
    }

    /// <summary>
    ///     Gets background information for a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <returns>A JSON string containing background information.</returns>
    private static string GetBackground(DocumentContext<Presentation> ctx, int slideIndex)
    {
        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var background = slide.Background;
        var fillFormat = background?.FillFormat;

        string? colorHex = null;
        double? opacity = null;

        if (fillFormat?.FillType == FillType.Solid)
            try
            {
                var solidColor = fillFormat.SolidFillColor.Color;
                if (!solidColor.IsEmpty)
                {
                    colorHex = solidColor.A < 255
                        ? $"#{solidColor.A:X2}{solidColor.R:X2}{solidColor.G:X2}{solidColor.B:X2}"
                        : $"#{solidColor.R:X2}{solidColor.G:X2}{solidColor.B:X2}";
                    opacity = Math.Round(solidColor.A / 255.0, 2);
                }
            }
            catch
            {
                // Theme colors may throw exceptions, return null for color
            }

        var result = new
        {
            slideIndex,
            hasBackground = background != null,
            fillType = fillFormat?.FillType.ToString(),
            color = colorHex,
            opacity,
            isPictureFill = fillFormat?.FillType == FillType.Picture
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}