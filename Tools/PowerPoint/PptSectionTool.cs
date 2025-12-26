using System.Text.Json;
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
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddSectionAsync(path, outputPath, arguments),
            "rename" => await RenameSectionAsync(path, outputPath, arguments),
            "delete" => await DeleteSectionAsync(path, outputPath, arguments),
            "get" => await GetSectionsAsync(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a section to the presentation
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sectionName, slideIndex</param>
    /// <returns>Success message</returns>
    private Task<string> AddSectionAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            presentation.Sections.AddSection(name, slide);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Section '{name}' added starting at slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Renames a section
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sectionIndex, newName</param>
    /// <returns>Success message</returns>
    private Task<string> RenameSectionAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex");
            var newName = ArgumentHelper.GetString(arguments, "newName");

            using var presentation = new Presentation(path);
            if (sectionIndex < 0 || sectionIndex >= presentation.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {presentation.Sections.Count - 1}");

            presentation.Sections[sectionIndex].Name = newName;
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Section {sectionIndex} renamed to '{newName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a section from the presentation
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteSectionAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex");
            var keepSlides = ArgumentHelper.GetBool(arguments, "keepSlides", true);

            using var presentation = new Presentation(path);
            PowerPointHelper.ValidateCollectionIndex(sectionIndex, presentation.Sections.Count, "section");
            var section = presentation.Sections[sectionIndex];
            if (keepSlides)
                presentation.Sections.RemoveSection(section);
            else
                presentation.Sections.RemoveSectionWithSlides(section);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Section {sectionIndex} removed (keep slides: {keepSlides}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all sections from the presentation
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>JSON string with all sections</returns>
    private Task<string> GetSectionsAsync(string path)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);

            if (presentation.Sections.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    sections = Array.Empty<object>(),
                    message = "No sections found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var sectionsList = new List<object>();
            for (var i = 0; i < presentation.Sections.Count; i++)
            {
                var sec = presentation.Sections[i];
                sectionsList.Add(new
                {
                    index = i,
                    name = sec.Name
                });
            }

            var result = new
            {
                count = presentation.Sections.Count,
                sections = sectionsList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}