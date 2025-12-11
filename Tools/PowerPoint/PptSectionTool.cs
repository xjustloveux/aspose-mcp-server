using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint sections (add, rename, delete, get)
/// Merges: PptAddSectionTool, PptRenameSectionTool, PptDeleteSectionTool, PptGetSectionsTool
/// </summary>
public class PptSectionTool : IAsposeTool
{
    public string Description => "Manage PowerPoint sections: add, rename, delete, or get";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'rename', 'delete', 'get'",
                @enum = new[] { "add", "rename", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            name = new
            {
                type = "string",
                description = "Section name (required for add/rename)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Start slide index for section (0-based, required for add)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, required for rename/delete)"
            },
            newName = new
            {
                type = "string",
                description = "New section name (required for rename)"
            },
            keepSlides = new
            {
                type = "boolean",
                description = "Keep slides in presentation (optional, for delete, default: true)"
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
            "add" => await AddSectionAsync(arguments, path),
            "rename" => await RenameSectionAsync(arguments, path),
            "delete" => await DeleteSectionAsync(arguments, path),
            "get" => await GetSectionsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddSectionAsync(JsonObject? arguments, string path)
    {
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required for add operation");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for add operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        presentation.Sections.AddSection(name, slide);
        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已新增章節 '{name}' 起始於投影片 {slideIndex}");
    }

    private async Task<string> RenameSectionAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? throw new ArgumentException("sectionIndex is required for rename operation");
        var newName = arguments?["newName"]?.GetValue<string>() ?? throw new ArgumentException("newName is required for rename operation");

        using var presentation = new Presentation(path);
        if (sectionIndex < 0 || sectionIndex >= presentation.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {presentation.Sections.Count - 1}");
        }

        presentation.Sections[sectionIndex].Name = newName;
        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已重新命名章節 {sectionIndex} 為 '{newName}'");
    }

    private async Task<string> DeleteSectionAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? throw new ArgumentException("sectionIndex is required for delete operation");
        var keepSlides = arguments?["keepSlides"]?.GetValue<bool?>() ?? true;

        using var presentation = new Presentation(path);
        PowerPointHelper.ValidateCollectionIndex(sectionIndex, presentation.Sections.Count, "章節");
        var section = presentation.Sections[sectionIndex];
        if (keepSlides)
        {
            presentation.Sections.RemoveSection(section);
        }
        else
        {
            presentation.Sections.RemoveSectionWithSlides(section);
        }
        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已移除章節 {sectionIndex}，保留投影片: {keepSlides}");
    }

    private async Task<string> GetSectionsAsync(JsonObject? arguments, string path)
    {
        using var presentation = new Presentation(path);
        var sb = new StringBuilder();
        sb.AppendLine($"Sections: {presentation.Sections.Count}");
        for (int i = 0; i < presentation.Sections.Count; i++)
        {
            var sec = presentation.Sections[i];
            sb.AppendLine($"[{i}] {sec.Name}");
        }
        return await Task.FromResult(sb.ToString());
    }
}

