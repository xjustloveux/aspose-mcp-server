using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint page setup (slide size, orientation, footer, slide numbering).
///     Merges: PptSetSlideSizeTool, PptSetSlideOrientationTool, PptHeaderFooterTool
/// </summary>
public class PptPageSetupTool : IAsposeTool
{
    private const float MinSizePoints = 1f;
    private const float MaxSizePoints = 5000f;

    public string Description =>
        @"Manage PowerPoint page setup. Supports 4 operations: set_size, set_orientation, set_footer, set_slide_numbering.

Note: PowerPoint slides do not have a separate header field. Only footer, date, and slide number are available.
Size unit: 1 inch = 72 points. Valid range: 1-5000 points.

Usage examples:
- Set slide size: ppt_page_setup(operation='set_size', path='presentation.pptx', preset='OnScreen16x9')
- Set custom size: ppt_page_setup(operation='set_size', path='presentation.pptx', preset='Custom', width=960, height=720)
- Set orientation: ppt_page_setup(operation='set_orientation', path='presentation.pptx', orientation='Portrait')
- Set footer: ppt_page_setup(operation='set_footer', path='presentation.pptx', footerText='Footer', showSlideNumber=true)
- Set slide numbering: ppt_page_setup(operation='set_slide_numbering', path='presentation.pptx', showSlideNumber=true, firstNumber=1)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set_size': Set slide size (required params: path, preset)
- 'set_orientation': Set slide orientation (required params: path, orientation)
- 'set_footer': Set footer text, date, slide number for slides (required params: path)
- 'set_slide_numbering': Set slide numbering visibility and start number (required params: path)",
                @enum = new[] { "set_size", "set_orientation", "set_footer", "set_slide_numbering" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            preset = new
            {
                type = "string",
                description = "Preset: OnScreen16x9, OnScreen16x10, Letter, A4, Banner, Custom (optional, for set_size)"
            },
            width = new
            {
                type = "number",
                description = "Custom width in points when preset=Custom (1-5000, 1 inch = 72 points)"
            },
            height = new
            {
                type = "number",
                description = "Custom height in points when preset=Custom (1-5000, 1 inch = 72 points)"
            },
            orientation = new
            {
                type = "string",
                description = "Orientation: 'Portrait' or 'Landscape' (required for set_orientation)",
                @enum = new[] { "Portrait", "Landscape" }
            },
            footerText = new
            {
                type = "string",
                description = "Footer text (optional, for set_footer)"
            },
            dateText = new
            {
                type = "string",
                description = "Date/time text (optional, for set_footer)"
            },
            showSlideNumber = new
            {
                type = "boolean",
                description = "Show slide number (optional, for set_footer/set_slide_numbering, default: true)"
            },
            firstNumber = new
            {
                type = "number",
                description = "First slide number (optional, for set_slide_numbering, default: 1)"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices (0-based, optional, for set_footer, if not provided applies to all slides)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "set_size" => await SetSlideSizeAsync(path, outputPath, arguments),
            "set_orientation" => await SetSlideOrientationAsync(path, outputPath, arguments),
            "set_footer" => await SetFooterAsync(path, outputPath, arguments),
            "set_slide_numbering" => await SetSlideNumberingAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets slide size.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing preset, width, height.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when custom size is missing width/height or values are out of range.</exception>
    private Task<string> SetSlideSizeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var preset = ArgumentHelper.GetString(arguments, "preset", "OnScreen16x9");
            var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
            var height = ArgumentHelper.GetDoubleNullable(arguments, "height");

            using var presentation = new Presentation(path);
            var slideSize = presentation.SlideSize;
            var type = preset.ToLower() switch
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

            presentation.Save(outputPath, SaveFormat.Pptx);
            var sizeInfo = slideSize.Type == SlideSizeType.Custom
                ? $" ({slideSize.Size.Width}x{slideSize.Size.Height})"
                : "";
            return $"Slide size set to {slideSize.Type}{sizeInfo}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Validates that width and height are within acceptable range.
    /// </summary>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    /// <exception cref="ArgumentException">Thrown when values are out of range (1-5000 points).</exception>
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
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing orientation (Portrait/Landscape).</param>
    /// <returns>Success message.</returns>
    private Task<string> SetSlideOrientationAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var orientation = ArgumentHelper.GetString(arguments, "orientation");
            var isPortrait = orientation.Equals("Portrait", StringComparison.OrdinalIgnoreCase);

            using var presentation = new Presentation(path);
            var currentSize = presentation.SlideSize.Size;
            var currentWidth = currentSize.Width;
            var currentHeight = currentSize.Height;

            var needsSwap = isPortrait ? currentWidth > currentHeight : currentHeight > currentWidth;

            if (needsSwap)
                presentation.SlideSize.SetSize(currentHeight, currentWidth, SlideSizeScaleType.EnsureFit);

            presentation.Save(outputPath, SaveFormat.Pptx);

            var finalSize = presentation.SlideSize.Size;
            return
                $"Slide orientation set to {orientation} ({finalSize.Width}x{finalSize.Height}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets footer text, date, and slide number for slides.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing footerText, dateText, showSlideNumber, slideIndices.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slide index is out of range.</exception>
    private Task<string> SetFooterAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var footerText = ArgumentHelper.GetStringNullable(arguments, "footerText");
            var showSlideNumber = ArgumentHelper.GetBool(arguments, "showSlideNumber", true);
            var dateText = ArgumentHelper.GetStringNullable(arguments, "dateText");
            var slideIndices = GetSlideIndices(arguments);

            using var presentation = new Presentation(path);

            var slides = GetTargetSlides(presentation, slideIndices);
            var applyToAll = slideIndices == null || slideIndices.Length == 0;

            if (applyToAll)
                EnableMasterVisibility(presentation, footerText, showSlideNumber, dateText);

            foreach (var slide in slides)
                ApplyFooterSettings(slide.HeaderFooterManager, footerText, showSlideNumber, dateText);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Footer settings updated for {slides.Count} slide(s). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets slide numbering visibility and start number.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing showSlideNumber and optional firstNumber.</param>
    /// <returns>Success message.</returns>
    private Task<string> SetSlideNumberingAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var showSlideNumber = ArgumentHelper.GetBool(arguments, "showSlideNumber", true);
            var firstNumber = ArgumentHelper.GetInt(arguments, "firstNumber", 1);

            using var presentation = new Presentation(path);

            presentation.FirstSlideNumber = firstNumber;
            presentation.HeaderFooterManager.SetAllSlideNumbersVisibility(showSlideNumber);

            foreach (var slide in presentation.Slides)
                slide.HeaderFooterManager.SetSlideNumberVisibility(showSlideNumber);

            presentation.Save(outputPath, SaveFormat.Pptx);

            var visibilityText = showSlideNumber ? "shown" : "hidden";
            return $"Slide numbers {visibilityText}, starting from {firstNumber}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets slide indices from arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional slideIndices.</param>
    /// <returns>Array of slide indices, or null if not specified.</returns>
    private static int[]? GetSlideIndices(JsonObject? arguments)
    {
        var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices", false);
        if (slideIndicesArray == null || slideIndicesArray.Count == 0)
            return null;

        return slideIndicesArray
            .Select(x => x?.GetValue<int>())
            .Where(x => x.HasValue)
            .Select(x => x!.Value)
            .ToArray();
    }

    /// <summary>
    ///     Gets target slides based on slide indices.
    /// </summary>
    /// <param name="presentation">Presentation to get slides from.</param>
    /// <param name="slideIndices">Slide indices, or null for all slides.</param>
    /// <returns>List of target slides.</returns>
    /// <exception cref="ArgumentException">Thrown when slide index is out of range.</exception>
    private static List<ISlide> GetTargetSlides(IPresentation presentation, int[]? slideIndices)
    {
        if (slideIndices == null || slideIndices.Length == 0)
            return presentation.Slides.ToList();

        var slides = new List<ISlide>();
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
    /// <param name="presentation">Presentation to configure.</param>
    /// <param name="footerText">Footer text (if not null, enables footer visibility).</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="dateText">Date text (if not null, enables date visibility).</param>
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
    /// <param name="manager">Slide header/footer manager.</param>
    /// <param name="footerText">Footer text.</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="dateText">Date text.</param>
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