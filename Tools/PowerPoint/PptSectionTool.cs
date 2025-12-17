using System.Text;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint sections (add, rename, delete, get)
///     Merges: PptAddSectionTool, PptRenameSectionTool, PptDeleteSectionTool, PptGetSectionsTool
/// </summary>
public class PptSectionTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint sections. Supports 4 operations: add, rename, delete, get.

Usage examples:
- Add section: ppt_section(operation='add', path='presentation.pptx', name='Section 1', slideIndex=0)
- Rename section: ppt_section(operation='rename', path='presentation.pptx', sectionIndex=0, newName='New Section')
- Delete section: ppt_section(operation='delete', path='presentation.pptx', sectionIndex=0)
- Get sections: ppt_section(operation='get', path='presentation.pptx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a new section (required params: path, name, slideIndex)
- 'rename': Rename a section (required params: path, sectionIndex, newName)
- 'delete': Delete a section (required params: path, sectionIndex)
- 'get': Get all sections (required params: path)",
                @enum = new[] { "add", "rename", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/delete/rename operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "add" => await AddSectionAsync(arguments, path),
            "rename" => await RenameSectionAsync(arguments, path),
            "delete" => await DeleteSectionAsync(arguments, path),
            "get" => await GetSectionsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a section to the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments containing sectionName, slideIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddSectionAsync(JsonObject? arguments, string path)
    {
        var name = ArgumentHelper.GetString(arguments, "name");
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        presentation.Sections.AddSection(name, slide);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Section '{name}' added starting at slide {slideIndex}");
    }

    /// <summary>
    ///     Renames a section
    /// </summary>
    /// <param name="arguments">JSON arguments containing sectionIndex, newName, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> RenameSectionAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex");
        var newName = ArgumentHelper.GetString(arguments, "newName");

        using var presentation = new Presentation(path);
        if (sectionIndex < 0 || sectionIndex >= presentation.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {presentation.Sections.Count - 1}");

        presentation.Sections[sectionIndex].Name = newName;
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Section {sectionIndex} renamed to '{newName}'");
    }

    /// <summary>
    ///     Deletes a section from the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments containing sectionIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteSectionAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex");
        var keepSlides = ArgumentHelper.GetBool(arguments, "keepSlides");

        using var presentation = new Presentation(path);
        PowerPointHelper.ValidateCollectionIndex(sectionIndex, presentation.Sections.Count, "section");
        var section = presentation.Sections[sectionIndex];
        if (keepSlides)
            presentation.Sections.RemoveSection(section);
        else
            presentation.Sections.RemoveSectionWithSlides(section);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Section {sectionIndex} removed, keep slides: {keepSlides}");
    }

    /// <summary>
    ///     Gets all sections from the presentation
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Formatted string with all sections</returns>
    private async Task<string> GetSectionsAsync(JsonObject? _, string path)
    {
        using var presentation = new Presentation(path);
        var sb = new StringBuilder();
        sb.AppendLine($"Sections: {presentation.Sections.Count}");
        for (var i = 0; i < presentation.Sections.Count; i++)
        {
            var sec = presentation.Sections[i];
            sb.AppendLine($"[{i}] {sec.Name}");
        }

        return await Task.FromResult(sb.ToString());
    }
}