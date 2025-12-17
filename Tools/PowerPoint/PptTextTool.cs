using System.Text;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint text (add, edit, replace)
///     Merges: PptAddTextTool, PptEditTextTool, PptReplaceTextTool
/// </summary>
public class PptTextTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint text. Supports 3 operations: add, edit, replace.

Usage examples:
- Add text: ppt_text(operation='add', path='presentation.pptx', slideIndex=0, text='Hello World', x=100, y=100)
- Edit text: ppt_text(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, text='Updated Text')
- Replace text: ppt_text(operation='replace', path='presentation.pptx', findText='old', replaceText='new')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add text to slide (required params: path, slideIndex, text)
- 'edit': Edit text in shape (required params: path, slideIndex, shapeIndex, text)
- 'replace': Replace text in presentation (required params: path, findText, replaceText)",
                @enum = new[] { "add", "edit", "replace" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for add/edit)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, required for edit)"
            },
            text = new
            {
                type = "string",
                description = "Text content (required for add/edit)"
            },
            findText = new
            {
                type = "string",
                description = "Text to find (required for replace)"
            },
            replaceText = new
            {
                type = "string",
                description = "Text to replace with (required for replace)"
            },
            matchCase = new
            {
                type = "boolean",
                description = "Match case (optional, for replace, default: false)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for add, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for add, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Text box width (optional, for add, default: 400)"
            },
            height = new
            {
                type = "number",
                description = "Text box height (optional, for add, default: 100)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
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
            "add" => await AddTextAsync(arguments, path),
            "edit" => await EditTextAsync(arguments, path),
            "replace" => await ReplaceTextAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds text to a slide
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing slideIndex, text, optional x, y, width, height, fontSize, fontName,
    ///     fontColor, outputPath
    /// </param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddTextAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var text = ArgumentHelper.GetString(arguments, "text");
        var x = ArgumentHelper.GetFloat(arguments, "x", "x", false, 50);
        var y = ArgumentHelper.GetFloat(arguments, "y", "y", false, 50);
        var width = ArgumentHelper.GetFloat(arguments, "width", "width", false, 400);
        var height = ArgumentHelper.GetFloat(arguments, "height", "height", false, 100);

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        var slide = presentation.Slides[slideIndex];
        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        textBox.TextFrame.Text = text;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);

        return await Task.FromResult($"Text added to slide {slideIndex}: {outputPath}");
    }

    /// <summary>
    ///     Edits text on a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, text, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditTextAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
        var text = ArgumentHelper.GetString(arguments, "text");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is IAutoShape { TextFrame: not null } autoShape)
            autoShape.TextFrame.Text = text;
        else
            throw new ArgumentException($"Shape at index {shapeIndex} does not support text editing");

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Text updated on slide {slideIndex}, shape {shapeIndex}");
    }

    /// <summary>
    ///     Replaces text in the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments containing searchText, replaceText, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message with replacement count</returns>
    private async Task<string> ReplaceTextAsync(JsonObject? arguments, string path)
    {
        var findText = ArgumentHelper.GetString(arguments, "findText");
        var replaceText = ArgumentHelper.GetString(arguments, "replaceText");
        var matchCase = ArgumentHelper.GetBool(arguments, "matchCase", false);

        using var presentation = new Presentation(path);
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var replacements = 0;

        foreach (var slide in presentation.Slides)
        foreach (var shape in slide.Shapes)
            if (shape is IAutoShape { TextFrame: not null } autoShape)
                foreach (var paragraph in autoShape.TextFrame.Paragraphs)
                foreach (var portion in paragraph.Portions)
                {
                    var text = portion.Text;
                    if (string.IsNullOrEmpty(text)) continue;

                    var newText = ReplaceAll(text, findText, replaceText, comparison);
                    if (!ReferenceEquals(text, newText) && newText != text)
                    {
                        portion.Text = newText;
                        replacements++;
                    }
                }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult(
            $"Text replacement completed: {replacements} occurrences\nFind: {findText}\nReplace with: {replaceText}\nOutput: {outputPath}");
    }

    private static string ReplaceAll(string source, string find, string replace, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(find)) return source;

        var sb = new StringBuilder();
        var idx = 0;
        while (true)
        {
            var next = source.IndexOf(find, idx, comparison);
            if (next < 0)
            {
                sb.Append(source, idx, source.Length - idx);
                break;
            }

            sb.Append(source, idx, next - idx);
            sb.Append(replace);
            idx = next + find.Length;
        }

        return sb.ToString();
    }
}