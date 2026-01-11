using System.Text.Json;
using Aspose.Slides;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Helper class for PowerPoint layout operations.
/// </summary>
public static class PptLayoutHelper
{
    /// <summary>
    ///     Mapping of layout type string names to SlideLayoutType enum values.
    /// </summary>
    private static readonly Dictionary<string, SlideLayoutType> LayoutTypeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["title"] = SlideLayoutType.Title,
        ["titleonly"] = SlideLayoutType.TitleOnly,
        ["blank"] = SlideLayoutType.Blank,
        ["twocolumn"] = SlideLayoutType.TwoColumnText,
        ["twocolumntext"] = SlideLayoutType.TwoColumnText,
        ["sectionheader"] = SlideLayoutType.SectionHeader,
        ["titleandcontent"] = SlideLayoutType.TitleAndObject,
        ["titleandobject"] = SlideLayoutType.TitleAndObject,
        ["objectandtext"] = SlideLayoutType.ObjectAndText,
        ["pictureandcaption"] = SlideLayoutType.PictureAndCaption
    };

    /// <summary>
    ///     Comma-separated list of supported layout type names for error messages.
    /// </summary>
    public static readonly string SupportedLayoutTypes = string.Join(", ", LayoutTypeMap.Keys);

    /// <summary>
    ///     Finds a layout slide by layout type string.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="layoutStr">The layout type string.</param>
    /// <returns>The matching layout slide.</returns>
    /// <exception cref="ArgumentException">Thrown when the layout type is unknown or not found in the presentation.</exception>
    public static ILayoutSlide FindLayoutByType(IPresentation presentation, string layoutStr)
    {
        if (!LayoutTypeMap.TryGetValue(layoutStr, out var layoutType))
            throw new ArgumentException(
                $"Unknown layout type: '{layoutStr}'. Supported types: {SupportedLayoutTypes}");

        var layout = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType);
        if (layout == null)
            throw new ArgumentException(
                $"Layout type '{layoutStr}' not found in this presentation. Use get_layouts to see available layouts.");

        return layout;
    }

    /// <summary>
    ///     Validates all slide indices before processing.
    /// </summary>
    /// <param name="indices">The array of slide indices to validate.</param>
    /// <param name="slideCount">The total number of slides in the presentation.</param>
    /// <exception cref="ArgumentException">Thrown when any slide index is out of range.</exception>
    public static void ValidateSlideIndices(int[] indices, int slideCount)
    {
        var invalidIndices = indices.Where(idx => idx < 0 || idx >= slideCount).ToList();
        if (invalidIndices.Count > 0)
            throw new ArgumentException(
                $"Invalid slide indices: [{string.Join(", ", invalidIndices)}]. Valid range: 0 to {slideCount - 1}");
    }

    /// <summary>
    ///     Builds a list of layout information including layout type.
    /// </summary>
    /// <param name="layoutSlides">The collection of layout slides.</param>
    /// <returns>A list of objects containing layout information.</returns>
    public static List<object> BuildLayoutsList(IMasterLayoutSlideCollection layoutSlides)
    {
        List<object> layoutsList = [];
        for (var i = 0; i < layoutSlides.Count; i++)
        {
            var layout = layoutSlides[i];
            layoutsList.Add(new
            {
                index = i,
                name = layout.Name,
                layoutType = layout.LayoutType.ToString()
            });
        }

        return layoutsList;
    }

    /// <summary>
    ///     Parses JSON array of slide indices.
    /// </summary>
    /// <param name="slideIndicesJson">The JSON string containing slide indices array.</param>
    /// <returns>An array of slide indices, or null if input is empty.</returns>
    /// <exception cref="ArgumentException">Thrown when the JSON format is invalid.</exception>
    public static int[]? ParseSlideIndicesJson(string? slideIndicesJson)
    {
        if (string.IsNullOrWhiteSpace(slideIndicesJson))
            return null;

        try
        {
            return JsonSerializer.Deserialize<int[]>(slideIndicesJson);
        }
        catch (JsonException)
        {
            throw new ArgumentException(
                $"Invalid slideIndices format. Expected JSON array, e.g., [0,1,2]. Got: {slideIndicesJson}");
        }
    }
}
