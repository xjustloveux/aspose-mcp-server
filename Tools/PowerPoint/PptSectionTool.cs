using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint sections (add, rename, delete, get).
/// </summary>
[McpServerToolType]
public class PptSectionTool
{
    /// <summary>
    ///     JSON serializer options for consistent output formatting.
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptSectionTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptSectionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint section operation (add, rename, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, rename, delete, get.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="name">Section name (required for add).</param>
    /// <param name="slideIndex">Start slide index for section (0-based, required for add).</param>
    /// <param name="sectionIndex">Section index (0-based, required for rename/delete).</param>
    /// <param name="newName">New section name (required for rename).</param>
    /// <param name="keepSlides">Keep slides in presentation (optional, for delete, default: true).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_section")]
    [Description(@"Manage PowerPoint sections. Supports 4 operations: add, rename, delete, get.

Warning: If outputPath is not provided for add/rename/delete operations, the original file will be overwritten.

Usage examples:
- Add section: ppt_section(operation='add', path='presentation.pptx', name='Section 1', slideIndex=0)
- Rename section: ppt_section(operation='rename', path='presentation.pptx', sectionIndex=0, newName='New Section')
- Delete section: ppt_section(operation='delete', path='presentation.pptx', sectionIndex=0)
- Get sections: ppt_section(operation='get', path='presentation.pptx')")]
    public string Execute(
        [Description("Operation: add, rename, delete, get")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Section name (required for add)")]
        string? name = null,
        [Description("Start slide index for section (0-based, required for add)")]
        int? slideIndex = null,
        [Description("Section index (0-based, required for rename/delete)")]
        int? sectionIndex = null,
        [Description("New section name (required for rename)")]
        string? newName = null,
        [Description("Keep slides in presentation (optional, for delete, default: true)")]
        bool keepSlides = true)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddSection(ctx, outputPath, name, slideIndex),
            "rename" => RenameSection(ctx, outputPath, sectionIndex, newName),
            "delete" => DeleteSection(ctx, outputPath, sectionIndex, keepSlides),
            "get" => GetSections(ctx),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a section to the presentation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="name">The section name.</param>
    /// <param name="slideIndex">The start slide index for the section.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when name or slideIndex is not provided.</exception>
    private static string AddSection(DocumentContext<Presentation> ctx, string? outputPath, string? name,
        int? slideIndex)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("name is required for add operation");
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for add operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        presentation.Sections.AddSection(name, slide);

        ctx.Save(outputPath);

        var result = $"Section '{name}' added starting at slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Renames a section.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="newName">The new section name.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex or newName is not provided.</exception>
    private static string RenameSection(DocumentContext<Presentation> ctx, string? outputPath, int? sectionIndex,
        string? newName)
    {
        if (!sectionIndex.HasValue)
            throw new ArgumentException("sectionIndex is required for rename operation");
        if (string.IsNullOrEmpty(newName))
            throw new ArgumentException("newName is required for rename operation");

        var presentation = ctx.Document;
        PowerPointHelper.ValidateCollectionIndex(sectionIndex.Value, presentation.Sections.Count, "section");

        presentation.Sections[sectionIndex.Value].Name = newName;

        ctx.Save(outputPath);

        var result = $"Section {sectionIndex} renamed to '{newName}'.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes a section from the presentation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="keepSlides">Whether to keep slides in the presentation.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is not provided.</exception>
    private static string DeleteSection(DocumentContext<Presentation> ctx, string? outputPath, int? sectionIndex,
        bool keepSlides)
    {
        if (!sectionIndex.HasValue)
            throw new ArgumentException("sectionIndex is required for delete operation");

        var presentation = ctx.Document;
        PowerPointHelper.ValidateCollectionIndex(sectionIndex.Value, presentation.Sections.Count, "section");
        var section = presentation.Sections[sectionIndex.Value];
        if (keepSlides)
            presentation.Sections.RemoveSection(section);
        else
            presentation.Sections.RemoveSectionWithSlides(section);

        ctx.Save(outputPath);

        var result = $"Section {sectionIndex} removed (keep slides: {keepSlides}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets all sections from the presentation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing section information.</returns>
    private static string GetSections(DocumentContext<Presentation> ctx)
    {
        var presentation = ctx.Document;

        if (presentation.Sections.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                sections = Array.Empty<object>(),
                message = "No sections found"
            };
            return JsonSerializer.Serialize(emptyResult, JsonOptions);
        }

        List<object> sectionsList = [];
        for (var i = 0; i < presentation.Sections.Count; i++)
        {
            var sec = presentation.Sections[i];
            var startSlideIndex = sec.StartedFromSlide != null
                ? presentation.Slides.IndexOf(sec.StartedFromSlide)
                : -1;
            sectionsList.Add(new
            {
                index = i,
                name = sec.Name,
                startSlideIndex,
                slideCount = sec.GetSlidesListOfSection().Count
            });
        }

        var result = new
        {
            count = presentation.Sections.Count,
            sections = sectionsList
        };

        return JsonSerializer.Serialize(result, JsonOptions);
    }
}