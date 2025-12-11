using System.Text.Json.Nodes;
using System.Text;
using System.Linq;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint layouts (set, get layouts, get masters, apply master, apply layout range, apply theme)
/// Merges: PptSetLayoutTool, PptGetLayoutsTool, PptGetMasterSlidesTool, PptApplyMasterTool, 
/// PptApplyLayoutRangeTool, PptApplyThemeTool
/// </summary>
public class PptLayoutTool : IAsposeTool
{
    public string Description => "Manage PowerPoint layouts: set, get layouts/masters, apply master/layout/theme";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'set', 'get_layouts', 'get_masters', 'apply_master', 'apply_layout_range', 'apply_theme'",
                @enum = new[] { "set", "get_layouts", "get_masters", "apply_master", "apply_layout_range", "apply_theme" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for set)"
            },
            layout = new
            {
                type = "string",
                description = "Layout type (Title, TitleOnly, Blank, TwoColumn, SectionHeader, etc., required for set/apply_layout_range)"
            },
            masterIndex = new
            {
                type = "number",
                description = "Master index (0-based, optional, for get_layouts/apply_master)"
            },
            layoutIndex = new
            {
                type = "number",
                description = "Layout index under master (0-based, optional, for apply_master)"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices array (required for apply_layout_range, optional for apply_master)"
            },
            themePath = new
            {
                type = "string",
                description = "Theme template file path (required for apply_theme)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        return operation.ToLower() switch
        {
            "set" => await SetLayoutAsync(arguments, path),
            "get_layouts" => await GetLayoutsAsync(arguments, path),
            "get_masters" => await GetMastersAsync(arguments, path),
            "apply_master" => await ApplyMasterAsync(arguments, path),
            "apply_layout_range" => await ApplyLayoutRangeAsync(arguments, path),
            "apply_theme" => await ApplyThemeAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> SetLayoutAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for set operation");
        var layoutStr = arguments?["layout"]?.GetValue<string>() ?? throw new ArgumentException("layout is required for set operation");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var layoutType = layoutStr.ToLower() switch
        {
            "title" => SlideLayoutType.Title,
            "titleonly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "twocolumn" => SlideLayoutType.TwoColumnText,
            "sectionheader" => SlideLayoutType.SectionHeader,
            _ => SlideLayoutType.Custom
        };

        var layout = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType) ?? presentation.LayoutSlides[0];
        presentation.Slides[slideIndex].LayoutSlide = layout;

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定投影片 {slideIndex} 版面：{layoutStr}");
    }

    private async Task<string> GetLayoutsAsync(JsonObject? arguments, string path)
    {
        var masterIndex = arguments?["masterIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        if (masterIndex.HasValue)
        {
            if (masterIndex.Value < 0 || masterIndex.Value >= presentation.Masters.Count)
            {
                throw new ArgumentException($"masterIndex must be between 0 and {presentation.Masters.Count - 1}");
            }
            var master = presentation.Masters[masterIndex.Value];
            sb.AppendLine($"=== Master {masterIndex.Value} Layouts ===");
            sb.AppendLine($"Total: {master.LayoutSlides.Count}");
            for (int i = 0; i < master.LayoutSlides.Count; i++)
            {
                var layout = master.LayoutSlides[i];
                sb.AppendLine($"  [{i}] {layout.Name ?? "(unnamed)"}");
            }
        }
        else
        {
            sb.AppendLine("=== All Layouts ===");
            for (int i = 0; i < presentation.Masters.Count; i++)
            {
                var master = presentation.Masters[i];
                sb.AppendLine($"\nMaster {i}: {master.LayoutSlides.Count} layout(s)");
                for (int j = 0; j < master.LayoutSlides.Count; j++)
                {
                    var layout = master.LayoutSlides[j];
                    sb.AppendLine($"  [{j}] {layout.Name ?? "(unnamed)"}");
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> GetMastersAsync(JsonObject? arguments, string path)
    {
        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        sb.AppendLine($"=== Master Slides ===");
        sb.AppendLine($"Total: {presentation.Masters.Count}");

        for (int i = 0; i < presentation.Masters.Count; i++)
        {
            var master = presentation.Masters[i];
            sb.AppendLine($"\nMaster {i}:");
            sb.AppendLine($"  Name: {master.Name ?? "(unnamed)"}");
            sb.AppendLine($"  Layouts: {master.LayoutSlides.Count}");
            for (int j = 0; j < master.LayoutSlides.Count; j++)
            {
                var layout = master.LayoutSlides[j];
                sb.AppendLine($"    [{j}] {layout.Name ?? "(unnamed)"}");
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> ApplyMasterAsync(JsonObject? arguments, string path)
    {
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var masterIndex = arguments?["masterIndex"]?.GetValue<int?>() ?? 0;
        var layoutIndex = arguments?["layoutIndex"]?.GetValue<int?>() ?? 0;

        using var presentation = new Presentation(path);

        PowerPointHelper.ValidateCollectionIndex(masterIndex, presentation.Masters.Count, "母版");
        var master = presentation.Masters[masterIndex];
        PowerPointHelper.ValidateCollectionIndex(layoutIndex, master.LayoutSlides.Count, "版面配置");

        var targets = slideIndices?.Length > 0
            ? slideIndices
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        foreach (var idx in targets)
        {
            if (idx < 0 || idx >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slide index {idx} out of range");
            }
        }

        var layout = master.LayoutSlides[layoutIndex];
        foreach (var idx in targets)
        {
            presentation.Slides[idx].LayoutSlide = layout;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已套用母片 {masterIndex} / 版面 {layoutIndex} 至 {targets.Length} 張投影片");
    }

    private async Task<string> ApplyLayoutRangeAsync(JsonObject? arguments, string path)
    {
        var layoutStr = arguments?["layout"]?.GetValue<string>() ?? throw new ArgumentException("layout is required for apply_layout_range operation");
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray()
                           ?? throw new ArgumentException("slideIndices is required for apply_layout_range operation");

        using var presentation = new Presentation(path);

        foreach (var idx in slideIndices)
        {
            if (idx < 0 || idx >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slide index {idx} out of range");
            }
        }

        var layoutType = layoutStr.ToLower() switch
        {
            "title" => SlideLayoutType.Title,
            "titleonly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "twocolumn" => SlideLayoutType.TwoColumnText,
            "sectionheader" => SlideLayoutType.SectionHeader,
            _ => SlideLayoutType.Custom
        };

        var layout = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType) ?? presentation.LayoutSlides[0];

        foreach (var idx in slideIndices)
        {
            presentation.Slides[idx].LayoutSlide = layout;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已套用版面 {layoutStr} 到 {slideIndices.Length} 張投影片");
    }

    private async Task<string> ApplyThemeAsync(JsonObject? arguments, string path)
    {
        var themePath = arguments?["themePath"]?.GetValue<string>() ?? throw new ArgumentException("themePath is required for apply_theme operation");

        using var presentation = new Presentation(path);
        using var themePresentation = new Presentation(themePath);

        // Copy theme from the first slide of theme presentation
        presentation.Slides[0].LayoutSlide = themePresentation.Slides[0].LayoutSlide;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Theme applied to presentation: {path}");
    }
}

