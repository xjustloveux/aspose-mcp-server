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
Searches text in AutoShapes, GroupShapes (recursive), and Table cells.

Coordinate unit: 1 inch = 72 points.

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
                description = "X position in points (optional, for add, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position in points (optional, for add, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Text box width in points (optional, for add, default: 400)"
            },
            height = new
            {
                type = "number",
                description = "Text box height in points (optional, for add, default: 100)"
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
            "add" => await AddTextAsync(path, outputPath, arguments),
            "edit" => await EditTextAsync(path, outputPath, arguments),
            "replace" => await ReplaceTextAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a text box to a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, text, optional x, y, width, height.</param>
    /// <returns>Success message.</returns>
    private Task<string> AddTextAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var text = ArgumentHelper.GetString(arguments, "text");
            var x = ArgumentHelper.GetFloat(arguments, "x", "x", false, 50);
            var y = ArgumentHelper.GetFloat(arguments, "y", "y", false, 50);
            var width = ArgumentHelper.GetFloat(arguments, "width", "width", false, 400);
            var height = ArgumentHelper.GetFloat(arguments, "height", "height", false, 100);

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
            textBox.TextFrame.Text = text;
            textBox.FillFormat.FillType = FillType.NoFill;
            textBox.LineFormat.FillFormat.FillType = FillType.NoFill;

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Text added to slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits text in a shape on a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, text.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not an AutoShape.</exception>
    private Task<string> EditTextAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var text = ArgumentHelper.GetString(arguments, "text");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            if (shape is not IAutoShape autoShape)
                throw new ArgumentException(
                    $"Shape at index {shapeIndex} (Type: {shape.GetType().Name}) is not an AutoShape and cannot contain text");

            if (autoShape.TextFrame == null)
                autoShape.AddTextFrame("");

            if (autoShape.TextFrame != null)
            {
                autoShape.TextFrame.Paragraphs.Clear();
                var paragraph = new Paragraph();
                paragraph.Portions.Add(new Portion(text));
                autoShape.TextFrame.Paragraphs.Add(paragraph);
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Text updated on slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Replaces text in the presentation across all shapes including GroupShapes and Tables.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing findText, replaceText, optional matchCase.</param>
    /// <returns>Success message with replacement count.</returns>
    private Task<string> ReplaceTextAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var findText = ArgumentHelper.GetString(arguments, "findText");
            var replaceText = ArgumentHelper.GetString(arguments, "replaceText");
            var matchCase = ArgumentHelper.GetBool(arguments, "matchCase", false);

            using var presentation = new Presentation(path);
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var replacements = 0;

            foreach (var slide in presentation.Slides)
                replacements += ProcessShapesForReplace(slide.Shapes, findText, replaceText, comparison);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Replaced '{findText}' with '{replaceText}' ({replacements} occurrences). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Recursively processes shapes for text replacement, including GroupShapes and Tables.
    /// </summary>
    /// <param name="shapes">Shape collection to process.</param>
    /// <param name="findText">Text to find.</param>
    /// <param name="replaceText">Text to replace with.</param>
    /// <param name="comparison">String comparison mode.</param>
    /// <returns>Number of replacements made.</returns>
    private static int ProcessShapesForReplace(IShapeCollection shapes, string findText, string replaceText,
        StringComparison comparison)
    {
        var replacements = 0;

        foreach (var shape in shapes)
            switch (shape)
            {
                case IAutoShape { TextFrame: not null } autoShape:
                    replacements += ReplaceInTextFrame(autoShape.TextFrame, findText, replaceText, comparison);
                    break;
                case IGroupShape groupShape:
                    replacements += ProcessShapesForReplace(groupShape.Shapes, findText, replaceText, comparison);
                    break;
                case ITable table:
                    replacements += ReplaceInTable(table, findText, replaceText, comparison);
                    break;
            }

        return replacements;
    }

    /// <summary>
    ///     Replaces text in a TextFrame while preserving formatting at the Portion level.
    /// </summary>
    /// <param name="textFrame">TextFrame to process.</param>
    /// <param name="findText">Text to find.</param>
    /// <param name="replaceText">Text to replace with.</param>
    /// <param name="comparison">String comparison mode.</param>
    /// <returns>1 if replacement was made, 0 otherwise.</returns>
    private static int ReplaceInTextFrame(ITextFrame textFrame, string findText, string replaceText,
        StringComparison comparison)
    {
        var originalText = textFrame.Text;
        if (string.IsNullOrEmpty(originalText)) return 0;

        if (originalText.IndexOf(findText, comparison) < 0) return 0;

        foreach (var para in textFrame.Paragraphs)
        foreach (var portion in para.Portions)
        {
            var portionText = portion.Text;
            if (string.IsNullOrEmpty(portionText)) continue;

            var newText = ReplaceAll(portionText, findText, replaceText, comparison);
            if (newText != portionText)
                portion.Text = newText;
        }

        return 1;
    }

    /// <summary>
    ///     Replaces text in all cells of a table.
    /// </summary>
    /// <param name="table">Table to process.</param>
    /// <param name="findText">Text to find.</param>
    /// <param name="replaceText">Text to replace with.</param>
    /// <param name="comparison">String comparison mode.</param>
    /// <returns>Number of cells where replacement was made.</returns>
    private static int ReplaceInTable(ITable table, string findText, string replaceText, StringComparison comparison)
    {
        var replacements = 0;

        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.TextFrame != null)
                replacements += ReplaceInTextFrame(cell.TextFrame, findText, replaceText, comparison);
        }

        return replacements;
    }

    /// <summary>
    ///     Replaces all occurrences of a string in source with replacement string.
    /// </summary>
    /// <param name="source">Source string.</param>
    /// <param name="find">Text to find.</param>
    /// <param name="replace">Text to replace with.</param>
    /// <param name="comparison">String comparison mode.</param>
    /// <returns>String with replacements made.</returns>
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