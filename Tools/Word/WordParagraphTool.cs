using System.Drawing;
using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for paragraph operations in Word documents
///     Merges: WordInsertParagraphTool, WordDeleteParagraphTool, WordEditParagraphTool,
///     WordGetParagraphsTool, WordGetParagraphFormatTool, WordCopyParagraphFormatTool, WordMergeParagraphsTool
/// </summary>
public class WordParagraphTool : IAsposeTool
{
    public string Description =>
        @"Manage paragraphs in Word documents. Supports 7 operations: insert, delete, edit, get, get_format, copy_format, merge.

Usage examples:
- Insert paragraph: word_paragraph(operation='insert', path='doc.docx', paragraphIndex=0, text='New paragraph')
- Delete paragraph: word_paragraph(operation='delete', path='doc.docx', paragraphIndex=0)
- Edit format: word_paragraph(operation='edit', path='doc.docx', paragraphIndex=0, alignment='center', fontSize=14)
- Get paragraph: word_paragraph(operation='get', path='doc.docx', paragraphIndex=0)
- Get format: word_paragraph(operation='get_format', path='doc.docx', paragraphIndex=0)
- Copy format: word_paragraph(operation='copy_format', path='doc.docx', sourceIndex=0, targetIndex=1)
- Merge paragraphs: word_paragraph(operation='merge', path='doc.docx', startIndex=0, endIndex=2)

Important notes for 'get' operation:
- By default, returns ALL paragraphs in the document structure, including paragraphs inside Comment objects
- Use includeCommentParagraphs=false to get only Body paragraphs (visible in document body)
- Each paragraph shows its ParentNode type to help identify its location
- Paragraphs inside Comment objects are marked with '[Comment]' in the location field";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'insert': Insert a new paragraph (required params: path, paragraphIndex, text)
- 'delete': Delete a paragraph (required params: path, paragraphIndex)
- 'edit': Edit paragraph format (required params: path, paragraphIndex)
- 'get': Get paragraph content (required params: path, paragraphIndex)
- 'get_format': Get paragraph format (required params: path, paragraphIndex)
- 'copy_format': Copy paragraph format (required params: path, sourceIndex, targetIndex)
- 'merge': Merge paragraphs (required params: path, startIndex, endIndex)",
                @enum = new[] { "insert", "delete", "edit", "get", "get_format", "copy_format", "merge" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for write operations)"
            },
            // Common parameters
            paragraphIndex = new
            {
                type = "number",
                description =
                    "Paragraph index (0-based, required for delete, edit, get_format operations, optional for insert/get operations). Valid range: 0 to (total paragraphs - 1), or -1 for last paragraph. Note: After delete operations, subsequent paragraph indices will shift automatically."
            },
            // Insert parameters
            text = new
            {
                type = "string",
                description = "Text content for the paragraph (required for insert operation)"
            },
            styleName = new
            {
                type = "string",
                description = "Style name to apply (e.g., 'Heading 1', 'Normal', optional, for insert/edit operations)"
            },
            alignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right, justify (optional, for insert/edit operations)",
                @enum = new[] { "left", "center", "right", "justify" }
            },
            // Get parameters
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, for get operation)"
            },
            includeEmpty = new
            {
                type = "boolean",
                description = "Include empty paragraphs (optional, default: true, for get operation)"
            },
            styleFilter = new
            {
                type = "string",
                description = "Filter by style name (optional, for get operation)"
            },
            includeCommentParagraphs = new
            {
                type = "boolean",
                description =
                    "Include paragraphs inside Comment objects (optional, default: true, for get operation). Set to false to get only Body paragraphs (visible in document body). Note: This returns the document's underlying structure, including Comment content that is not visible in the body."
            },
            includeTextboxParagraphs = new
            {
                type = "boolean",
                description =
                    "Include paragraphs inside TextBox/Shape objects (optional, default: true, for get operation). Set to false to exclude textbox paragraphs. Note: Textbox paragraphs are marked with [TextBox] in the location field."
            },
            // Get format parameters
            includeRunDetails = new
            {
                type = "boolean",
                description =
                    "Include detailed run-level formatting (optional, default: true, for get_format operation)"
            },
            // Edit parameters
            fontName = new
            {
                type = "string",
                description = "Font name (e.g., 'Arial', optional, for edit operation)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, optional, for edit operation)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description =
                    "Font name for Far East characters (Chinese/Japanese/Korean, optional, for edit operation)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (optional, for edit operation)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (optional, for edit operation)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (optional, for edit operation)"
            },
            underline = new
            {
                type = "boolean",
                description = "Underline text (optional, for edit operation)"
            },
            color = new
            {
                type = "string",
                description = "Font color (hex format, e.g., '000000' for black, optional, for edit operation)"
            },
            indentLeft = new
            {
                type = "number",
                description = "Left indent in points (optional, for insert/edit operations)"
            },
            indentRight = new
            {
                type = "number",
                description = "Right indent in points (optional, for insert/edit operations)"
            },
            firstLineIndent = new
            {
                type = "number",
                description =
                    "First line indent in points (positive for indent, negative for hanging, optional, for insert/edit operations)"
            },
            spaceBefore = new
            {
                type = "number",
                description = "Space before paragraph in points (optional, for insert/edit operations)"
            },
            spaceAfter = new
            {
                type = "number",
                description = "Space after paragraph in points (optional, for insert/edit operations)"
            },
            lineSpacing = new
            {
                type = "number",
                description =
                    "Line spacing (points or multiplier depending on lineSpacingRule, optional, for edit operation)"
            },
            lineSpacingRule = new
            {
                type = "string",
                description =
                    "Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple (optional, for edit operation)",
                @enum = new[] { "single", "oneAndHalf", "double", "atLeast", "exactly", "multiple" }
            },
            tabStops = new
            {
                type = "array",
                description = "Custom tab stops (optional, for edit operation)",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number" },
                        alignment = new
                            { type = "string", @enum = new[] { "left", "center", "right", "decimal", "bar", "clear" } },
                        leader = new
                        {
                            type = "string", @enum = new[] { "none", "dots", "dashes", "line", "heavy", "middleDot" }
                        }
                    }
                }
            },
            // Copy format parameters
            sourceParagraphIndex = new
            {
                type = "number",
                description = "Source paragraph index (0-based, required for copy_format operation)"
            },
            targetParagraphIndex = new
            {
                type = "number",
                description = "Target paragraph index (0-based, required for copy_format operation)"
            },
            // Merge parameters
            startParagraphIndex = new
            {
                type = "number",
                description = "Start paragraph index (0-based, inclusive, required for merge operation)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description = "End paragraph index (0-based, inclusive, required for merge operation)"
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
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        return operation switch
        {
            "insert" => await InsertParagraphAsync(arguments, path, outputPath),
            "delete" => await DeleteParagraphAsync(arguments, path, outputPath),
            "edit" => await EditParagraphAsync(arguments, path, outputPath),
            "get" => await GetParagraphsAsync(arguments, path),
            "get_format" => await GetParagraphFormatAsync(arguments, path),
            "copy_format" => await CopyParagraphFormatAsync(arguments, path, outputPath),
            "merge" => await MergeParagraphsAsync(arguments, path, outputPath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a paragraph into the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing text, optional paragraphIndex, styleName, formatting options</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> InsertParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = ArgumentHelper.GetString(arguments, "text");
        var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
        var styleName = ArgumentHelper.GetStringNullable(arguments, "styleName");
        var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment");
        var indentLeft = ArgumentHelper.GetDoubleNullable(arguments, "indentLeft");
        var indentRight = ArgumentHelper.GetDoubleNullable(arguments, "indentRight");
        var firstLineIndent = ArgumentHelper.GetDoubleNullable(arguments, "firstLineIndent");
        var spaceBefore = ArgumentHelper.GetDoubleNullable(arguments, "spaceBefore");
        var spaceAfter = ArgumentHelper.GetDoubleNullable(arguments, "spaceAfter");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        Paragraph? targetPara = null;
        var insertPosition = "end of document";

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[0] as Paragraph;
                    insertPosition = "beginning of document";
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                insertPosition = $"after paragraph #{paragraphIndex.Value}";
            }
            else
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
        }

        var para = new Paragraph(doc);
        var run = new Run(doc, text);
        para.AppendChild(run);

        if (!string.IsNullOrEmpty(styleName))
            try
            {
                var style = doc.Styles[styleName];
                if (style != null)
                    para.ParagraphFormat.StyleName = styleName;
                else
                    throw new ArgumentException(
                        $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Unable to apply style '{styleName}': {ex.Message}. Use word_get_styles tool to view available styles",
                    ex);
            }

        if (!string.IsNullOrEmpty(alignment))
            para.ParagraphFormat.Alignment = alignment.ToLower() switch
            {
                "left" => ParagraphAlignment.Left,
                "right" => ParagraphAlignment.Right,
                "center" => ParagraphAlignment.Center,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };

        // Apply indentation and spacing
        if (indentLeft.HasValue) para.ParagraphFormat.LeftIndent = indentLeft.Value;

        if (indentRight.HasValue) para.ParagraphFormat.RightIndent = indentRight.Value;

        if (firstLineIndent.HasValue) para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;

        if (spaceBefore.HasValue) para.ParagraphFormat.SpaceBefore = spaceBefore.Value;

        if (spaceAfter.HasValue) para.ParagraphFormat.SpaceAfter = spaceAfter.Value;

        if (targetPara != null)
        {
            if (paragraphIndex!.Value == -1)
                targetPara.ParentNode.InsertBefore(para, targetPara);
            else
                targetPara.ParentNode.InsertAfter(para, targetPara);
        }
        else
        {
            var body = doc.FirstSection.Body;
            body.AppendChild(para);
        }

        doc.Save(outputPath);

        var result = "Paragraph inserted successfully\n";
        result += $"Insert position: {insertPosition}\n";
        if (!string.IsNullOrEmpty(styleName)) result += $"Applied style: {styleName}\n";
        if (!string.IsNullOrEmpty(alignment)) result += $"Alignment: {alignment}\n";
        if (indentLeft.HasValue || indentRight.HasValue || firstLineIndent.HasValue)
        {
            result += "Indent: ";
            var indentParts = new List<string>();
            if (indentLeft.HasValue) indentParts.Add($"Left={indentLeft.Value}pt");
            if (indentRight.HasValue) indentParts.Add($"Right={indentRight.Value}pt");
            if (firstLineIndent.HasValue) indentParts.Add($"First line={firstLineIndent.Value}pt");
            result += string.Join(", ", indentParts) + "\n";
        }

        if (spaceBefore.HasValue || spaceAfter.HasValue)
        {
            result += "Spacing: ";
            var spaceParts = new List<string>();
            if (spaceBefore.HasValue) spaceParts.Add($"Before={spaceBefore.Value}pt");
            if (spaceAfter.HasValue) spaceParts.Add($"After={spaceAfter.Value}pt");
            result += string.Join(", ", spaceParts) + "\n";
        }

        result += $"Document paragraph count: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Deletes a paragraph from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Handle paragraphIndex=-1 (delete last paragraph)
        if (paragraphIndex == -1)
        {
            if (paragraphs.Count == 0)
                throw new ArgumentException("Cannot delete paragraph: document has no paragraphs");
            paragraphIndex = paragraphs.Count - 1;
        }

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        var paragraphToDelete = paragraphs[paragraphIndex] as Paragraph;
        if (paragraphToDelete == null)
            throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex}");

        var textPreview = paragraphToDelete.GetText().Trim();
        if (textPreview.Length > 50) textPreview = textPreview.Substring(0, 50) + "...";

        paragraphToDelete.Remove();

        doc.Save(outputPath);

        var result = $"Paragraph #{paragraphIndex} deleted successfully\n";
        if (!string.IsNullOrEmpty(textPreview)) result += $"Content preview: {textPreview}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Edits paragraph properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional text, formatting options, sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

        var doc = new Document(path);

        // Handle paragraphIndex=-1 (document end)
        if (paragraphIndex == -1)
        {
            var lastSection = doc.LastSection;
            var bodyParagraphs = lastSection.Body.GetChildNodes(NodeType.Paragraph, false);
            if (bodyParagraphs.Count > 0)
            {
                paragraphIndex = bodyParagraphs.Count - 1;
                sectionIndex = doc.Sections.Count - 1;
            }
            else
            {
                throw new ArgumentException(
                    "Cannot edit paragraph: document has no paragraphs. Use insert operation to add paragraphs first.");
            }
        }

        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException(
                $"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count}, valid range: 0-{doc.Sections.Count - 1})");

        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count}, valid range: 0-{paragraphs.Count - 1})");

        var para = paragraphs[paragraphIndex];
        var builder = new DocumentBuilder(doc);

        // Move to first child if exists, otherwise move to paragraph itself
        builder.MoveTo(para.FirstChild ?? para);

        // Apply font properties using FontHelper
        var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
        var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
        var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
        var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
        var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
        var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
        var underline = ArgumentHelper.GetBoolNullable(arguments, "underline");
        var colorStr = ArgumentHelper.GetStringNullable(arguments, "color");
        var underlineStr = underline.HasValue ? underline.Value ? "single" : "none" : null;

        FontHelper.Word.ApplyFontSettings(
            builder,
            fontName,
            fontNameAscii,
            fontNameFarEast,
            fontSize,
            bold,
            italic,
            underlineStr,
            colorStr
        );

        // Apply paragraph properties
        var paraFormat = para.ParagraphFormat;

        var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment") ?? "left";
        if (!string.IsNullOrEmpty(alignment))
            paraFormat.Alignment = GetAlignment(alignment);

        var indentLeft = ArgumentHelper.GetDoubleNullable(arguments, "indentLeft");
        if (indentLeft.HasValue)
            paraFormat.LeftIndent = indentLeft.Value;

        var indentRight = ArgumentHelper.GetDoubleNullable(arguments, "indentRight");
        if (indentRight.HasValue)
            paraFormat.RightIndent = indentRight.Value;

        var firstLineIndent = ArgumentHelper.GetDoubleNullable(arguments, "firstLineIndent");
        if (firstLineIndent.HasValue)
            paraFormat.FirstLineIndent = firstLineIndent.Value;

        var spaceBefore = ArgumentHelper.GetDoubleNullable(arguments, "spaceBefore");
        if (spaceBefore.HasValue)
            paraFormat.SpaceBefore = spaceBefore.Value;

        var spaceAfter = ArgumentHelper.GetDoubleNullable(arguments, "spaceAfter");
        if (spaceAfter.HasValue)
            paraFormat.SpaceAfter = spaceAfter.Value;

        if (arguments?["lineSpacing"] != null || arguments?["lineSpacingRule"] != null)
        {
            var lineSpacing = ArgumentHelper.GetDoubleNullable(arguments, "lineSpacing");
            var lineSpacingRule = ArgumentHelper.GetString(arguments, "lineSpacingRule", "single");

            var rule = GetLineSpacingRule(lineSpacingRule);
            paraFormat.LineSpacingRule = rule;

            if (lineSpacing.HasValue)
                paraFormat.LineSpacing = lineSpacing.Value;
            else if (lineSpacingRule == "single")
                paraFormat.LineSpacing = 12;
            else if (lineSpacingRule == "oneAndHalf")
                paraFormat.LineSpacing = 18;
            else if (lineSpacingRule == "double") paraFormat.LineSpacing = 24;
        }

        var styleName = ArgumentHelper.GetStringNullable(arguments, "styleName");
        if (!string.IsNullOrEmpty(styleName))
            try
            {
                paraFormat.Style = doc.Styles[styleName];
            }
            catch
            {
                // Style not found, ignore
            }

        // Apply tab stops
        var tabStops = ArgumentHelper.GetArray(arguments, "tabStops", false);
        if (tabStops is { Count: > 0 })
        {
            paraFormat.TabStops.Clear();
            foreach (var ts in tabStops)
            {
                var tsObj = ts?.AsObject();
                if (tsObj != null)
                {
                    var position = tsObj["position"]?.GetValue<double>() ?? 0;
                    var tabAlignment = tsObj["alignment"]?.GetValue<string>() ?? "left";
                    var leader = tsObj["leader"]?.GetValue<string>() ?? "none";

                    paraFormat.TabStops.Add(new TabStop(
                        position,
                        GetTabAlignment(tabAlignment),
                        GetTabLeader(leader)
                    ));
                }
            }
        }

        var textParam = ArgumentHelper.GetStringNullable(arguments, "text");
        if (!string.IsNullOrEmpty(textParam))
        {
            // Clear existing content and add new text
            para.RemoveAllChildren();
            var newRun = new Run(doc, textParam);

            // Apply font settings to the new run using FontHelper (reuse underlineStr from outer scope)
            FontHelper.Word.ApplyFontSettings(
                newRun,
                fontName,
                fontNameAscii,
                fontNameFarEast,
                fontSize,
                bold,
                italic,
                underlineStr,
                colorStr
            );

            para.AppendChild(newRun);
        }
        else
        {
            // Apply font to all existing runs in paragraph
            var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            if (runs.Count == 0)
            {
                // Paragraph has no runs, but user didn't provide text - this is OK, just apply formatting
                // Formatting will be applied when text is added later
            }
            else
            {
                foreach (var run in runs)
                    // Apply font settings using FontHelper (reuse underlineStr from outer scope)
                    FontHelper.Word.ApplyFontSettings(
                        run,
                        fontName,
                        fontNameAscii,
                        fontNameFarEast,
                        fontSize,
                        bold,
                        italic,
                        underlineStr,
                        colorStr
                    );
            }
        }

        doc.Save(outputPath);

        var resultMsg = $"Paragraph {paragraphIndex} format edited successfully";
        if (!string.IsNullOrEmpty(textParam))
            resultMsg += ", text content updated";
        return await Task.FromResult(resultMsg);
    }

    /// <summary>
    ///     Gets all paragraphs from the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing optional sectionIndex, includeCommentParagraphs,
    ///     includeTextboxParagraphs
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all paragraphs</returns>
    private async Task<string> GetParagraphsAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
        var includeEmpty = ArgumentHelper.GetBool(arguments, "includeEmpty");
        var styleFilter = ArgumentHelper.GetStringNullable(arguments, "styleFilter");
        var includeCommentParagraphs = ArgumentHelper.GetBool(arguments, "includeCommentParagraphs");
        var includeTextboxParagraphs = ArgumentHelper.GetBool(arguments, "includeTextboxParagraphs");

        var doc = new Document(path);
        var sb = new StringBuilder();

        List<Paragraph> paragraphs;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            // When sectionIndex is specified, only get paragraphs from that section's Body
            paragraphs = doc.Sections[sectionIndex.Value].Body
                .GetChildNodes(NodeType.Paragraph, includeCommentParagraphs).Cast<Paragraph>().ToList();
        }
        else
        {
            if (includeCommentParagraphs)
            {
                // Get all paragraphs including those inside Comment objects
                paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            }
            else
            {
                // Get only paragraphs from document Body (visible in document body)
                paragraphs = new List<Paragraph>();
                foreach (var section in doc.Sections.Cast<Section>())
                {
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
                        .ToList();
                    paragraphs.AddRange(bodyParagraphs);
                }
            }
        }

        if (!includeEmpty) paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.GetText())).ToList();

        if (!string.IsNullOrEmpty(styleFilter))
            paragraphs = paragraphs.Where(p => p.ParagraphFormat.Style?.Name == styleFilter).ToList();

        // Filter out textbox paragraphs if includeTextboxParagraphs is false
        if (!includeTextboxParagraphs)
            paragraphs = paragraphs.Where(p =>
            {
                // Check if paragraph is inside a Shape (including nested structures)
                var shapeAncestor = p.GetAncestor(NodeType.Shape);
                if (shapeAncestor is Shape { ShapeType: ShapeType.TextBox }) return false; // Exclude textbox paragraphs

                // Also check if paragraph's parent node chain contains a Shape
                // This handles cases where textbox content might be structured differently
                var currentNode = p.ParentNode;
                while (currentNode != null)
                {
                    if (currentNode.NodeType == NodeType.Shape)
                        if (currentNode is Shape { ShapeType: ShapeType.TextBox })
                            return false; // Exclude textbox paragraphs

                    currentNode = currentNode.ParentNode;
                }

                return true;
            }).ToList();

        sb.AppendLine($"=== Paragraphs ({paragraphs.Count}) ===");
        if (!includeCommentParagraphs && !includeTextboxParagraphs)
            sb.AppendLine(
                "(Only Body paragraphs - visible in document body, excluding Comment and TextBox paragraphs)");
        else if (!includeCommentParagraphs)
            sb.AppendLine("(Only Body paragraphs - visible in document body, excluding Comment paragraphs)");
        else if (!includeTextboxParagraphs)
            sb.AppendLine("(Includes all paragraphs in document structure, excluding TextBox paragraphs)");
        else
            sb.AppendLine("(Includes all paragraphs in document structure, including Comment objects and TextBoxes)");
        sb.AppendLine();

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();

            // Determine paragraph location
            var location = "Body";
            var parentNodeInfo = "";

            if (para.ParentNode != null)
            {
                var parentNodeType = para.ParentNode.NodeType;
                parentNodeInfo = $"ParentNode: {parentNodeType}";

                // Check if paragraph is inside a Comment object
                var commentAncestor = para.GetAncestor(NodeType.Comment);
                if (commentAncestor != null)
                {
                    location = "[Comment]";
                    if (commentAncestor is Comment comment)
                        parentNodeInfo += $" (Comment ID: {comment.Id}, Author: {comment.Author})";
                }
                else
                {
                    // Check if paragraph is inside a TextBox/Shape
                    var shapeAncestor = para.GetAncestor(NodeType.Shape);
                    if (shapeAncestor != null)
                    {
                        if (shapeAncestor is Shape { ShapeType: ShapeType.TextBox })
                        {
                            location = "[TextBox]";
                            parentNodeInfo += " (TextBox)";
                        }
                        else
                        {
                            location = "[Shape]";
                        }
                    }
                    else
                    {
                        // Check if paragraph is directly in Body
                        var bodyAncestor = para.GetAncestor(NodeType.Body);
                        if (bodyAncestor != null && para.ParentNode.NodeType == NodeType.Body)
                            location = "Body";
                        else if (para.ParentNode.NodeType == NodeType.Body)
                            location = "Body";
                        else
                            location = $"[{parentNodeType}]";
                    }
                }
            }

            sb.AppendLine($"[{i}] Location: {location}");
            sb.AppendLine($"    Style: {para.ParagraphFormat.Style?.Name ?? "(none)"}");
            if (!string.IsNullOrEmpty(parentNodeInfo)) sb.AppendLine($"    {parentNodeInfo}");
            sb.AppendLine(
                $"    Text: {text.Substring(0, Math.Min(100, text.Length))}{(text.Length > 100 ? "..." : "")}");
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    ///     Gets format information for a paragraph
    /// </summary>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with paragraph format details</returns>
    private async Task<string> GetParagraphFormatAsync(JsonObject? arguments, string path)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
        var includeRunDetails = ArgumentHelper.GetBool(arguments, "includeRunDetails", true);

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null) throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex}");

        var result = new StringBuilder();
        result.AppendLine($"=== Paragraph #{paragraphIndex} Format Information ===\n");

        result.AppendLine("[Basic Information]");
        result.AppendLine($"Paragraph text: {para.GetText().Trim()}");
        result.AppendLine($"Text length: {para.GetText().Trim().Length} characters");
        result.AppendLine($"Run count: {para.Runs.Count}");
        result.AppendLine();

        var format = para.ParagraphFormat;
        result.AppendLine("[Paragraph Format]");
        result.AppendLine($"Style name: {format.StyleName}");
        result.AppendLine($"Alignment: {format.Alignment}");
        result.AppendLine($"Left indent: {format.LeftIndent:F2} pt ({format.LeftIndent / 28.35:F2} cm)");
        result.AppendLine($"Right indent: {format.RightIndent:F2} pt ({format.RightIndent / 28.35:F2} cm)");
        result.AppendLine(
            $"First line indent: {format.FirstLineIndent:F2} pt ({format.FirstLineIndent / 28.35:F2} cm)");
        result.AppendLine($"Space before: {format.SpaceBefore:F2} pt");
        result.AppendLine($"Space after: {format.SpaceAfter:F2} pt");
        result.AppendLine($"Line spacing: {format.LineSpacing:F2} pt");
        result.AppendLine($"Line spacing rule: {format.LineSpacingRule}");
        result.AppendLine();

        if (para.ListFormat is { IsListItem: true })
        {
            result.AppendLine("[List Format]");
            result.AppendLine("Is list item: Yes");
            result.AppendLine($"List level: {para.ListFormat.ListLevelNumber}");
            if (para.ListFormat.List != null) result.AppendLine($"List ID: {para.ListFormat.List.ListId}");
            result.AppendLine();
        }

        if (format.Borders.Count > 0)
        {
            result.AppendLine("[Borders]");
            if (format.Borders.Top.LineStyle != LineStyle.None)
                result.AppendLine(
                    $"Top border: {format.Borders.Top.LineStyle}, {format.Borders.Top.LineWidth} pt, Color: {format.Borders.Top.Color.Name}");
            if (format.Borders.Bottom.LineStyle != LineStyle.None)
                result.AppendLine(
                    $"Bottom border: {format.Borders.Bottom.LineStyle}, {format.Borders.Bottom.LineWidth} pt, Color: {format.Borders.Bottom.Color.Name}");
            if (format.Borders.Left.LineStyle != LineStyle.None)
                result.AppendLine(
                    $"Left border: {format.Borders.Left.LineStyle}, {format.Borders.Left.LineWidth} pt, Color: {format.Borders.Left.Color.Name}");
            if (format.Borders.Right.LineStyle != LineStyle.None)
                result.AppendLine(
                    $"Right border: {format.Borders.Right.LineStyle}, {format.Borders.Right.LineWidth} pt, Color: {format.Borders.Right.Color.Name}");
            result.AppendLine();
        }

        if (format.Shading.BackgroundPatternColor.ToArgb() != Color.Empty.ToArgb())
        {
            result.AppendLine("[Background Color]");
            var color = format.Shading.BackgroundPatternColor;
            result.AppendLine($"Background color: #{color.R:X2}{color.G:X2}{color.B:X2}");
            result.AppendLine();
        }

        if (format.TabStops.Count > 0)
        {
            result.AppendLine("[Tab Stops]");
            for (var i = 0; i < format.TabStops.Count; i++)
            {
                var tab = format.TabStops[i];
                result.AppendLine(
                    $"  Tab {i + 1}: Position={tab.Position:F2} pt, Alignment={tab.Alignment}, Leader={tab.Leader}");
            }

            result.AppendLine();
        }

        if (para.Runs.Count > 0)
        {
            var firstRun = para.Runs[0];
            result.AppendLine("[Font Format (First Run)]");

            if (firstRun.Font.NameAscii != firstRun.Font.NameFarEast)
            {
                result.AppendLine($"Font (ASCII): {firstRun.Font.NameAscii}");
                result.AppendLine($"Font (Far East): {firstRun.Font.NameFarEast}");
            }
            else
            {
                result.AppendLine($"Font: {firstRun.Font.Name}");
            }

            result.AppendLine($"Font size: {firstRun.Font.Size} pt");

            if (firstRun.Font.Bold) result.AppendLine("Bold: Yes");
            if (firstRun.Font.Italic) result.AppendLine("Italic: Yes");
            if (firstRun.Font.Underline != Underline.None) result.AppendLine($"Underline: {firstRun.Font.Underline}");
            if (firstRun.Font.StrikeThrough) result.AppendLine("Strikethrough: Yes");
            if (firstRun.Font.Superscript) result.AppendLine("Superscript: Yes");
            if (firstRun.Font.Subscript) result.AppendLine("Subscript: Yes");

            if (firstRun.Font.Color.ToArgb() != Color.Empty.ToArgb())
            {
                var color = firstRun.Font.Color;
                result.AppendLine($"Color: #{color.R:X2}{color.G:X2}{color.B:X2}");
            }

            if (firstRun.Font.HighlightColor != Color.Empty)
                result.AppendLine($"Highlight: {firstRun.Font.HighlightColor.Name}");
            result.AppendLine();
        }

        if (includeRunDetails && para.Runs.Count > 1)
        {
            result.AppendLine("[Run Details]");
            result.AppendLine($"Total {para.Runs.Count} Runs:");

            for (var i = 0; i < Math.Min(para.Runs.Count, 10); i++)
            {
                var run = para.Runs[i];
                result.AppendLine($"\n  Run #{i}:");
                result.AppendLine($"    Text: {run.Text.Replace("\r", "\\r").Replace("\n", "\\n")}");

                if (run.Font.NameAscii != run.Font.NameFarEast)
                {
                    result.AppendLine($"    Font (ASCII): {run.Font.NameAscii}");
                    result.AppendLine($"    Font (Far East): {run.Font.NameFarEast}");
                }
                else
                {
                    result.AppendLine($"    Font: {run.Font.Name}");
                }

                result.AppendLine($"    Font size: {run.Font.Size} pt");

                var styles = new List<string>();
                if (run.Font.Bold) styles.Add("Bold");
                if (run.Font.Italic) styles.Add("Italic");
                if (run.Font.Underline != Underline.None) styles.Add($"Underline({run.Font.Underline})");
                if (styles.Count > 0)
                    result.AppendLine($"    Styles: {string.Join(", ", styles)}");
            }

            if (para.Runs.Count > 10) result.AppendLine($"\n  ... {para.Runs.Count - 10} more Runs (omitted)");
            result.AppendLine();
        }

        result.AppendLine("[JSON Format (for word_edit_paragraph)]");
        result.AppendLine("{");
        result.AppendLine($"  \"alignment\": \"{format.Alignment.ToString().ToLower()}\",");
        result.AppendLine($"  \"leftIndent\": {format.LeftIndent:F2},");
        result.AppendLine($"  \"rightIndent\": {format.RightIndent:F2},");
        result.AppendLine($"  \"firstLineIndent\": {format.FirstLineIndent:F2},");
        result.AppendLine($"  \"spaceBefore\": {format.SpaceBefore:F2},");
        result.AppendLine($"  \"spaceAfter\": {format.SpaceAfter:F2},");
        result.AppendLine($"  \"lineSpacing\": {format.LineSpacing:F2}");

        if (para.Runs.Count > 0)
        {
            var run = para.Runs[0];
            result.AppendLine($"  \"fontNameAscii\": \"{run.Font.NameAscii}\",");
            result.AppendLine($"  \"fontNameFarEast\": \"{run.Font.NameFarEast}\",");
            result.AppendLine($"  \"fontSize\": {run.Font.Size},");
            result.AppendLine($"  \"bold\": {run.Font.Bold.ToString().ToLower()},");
            result.AppendLine($"  \"italic\": {run.Font.Italic.ToString().ToLower()}");
        }

        result.AppendLine("}");

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    ///     Copies paragraph format from source to target
    /// </summary>
    /// <param name="arguments">JSON arguments containing sourceIndex, targetIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> CopyParagraphFormatAsync(JsonObject? arguments, string path, string outputPath)
    {
        var sourceParagraphIndex = ArgumentHelper.GetInt(arguments, "sourceParagraphIndex");
        var targetParagraphIndex = ArgumentHelper.GetInt(arguments, "targetParagraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (sourceParagraphIndex < 0 || sourceParagraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Source paragraph index {sourceParagraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (targetParagraphIndex < 0 || targetParagraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Target paragraph index {targetParagraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        var sourcePara = paragraphs[sourceParagraphIndex] as Paragraph;
        var targetPara = paragraphs[targetParagraphIndex] as Paragraph;

        if (sourcePara == null || targetPara == null) throw new InvalidOperationException("Unable to get paragraphs");

        targetPara.ParagraphFormat.StyleName = sourcePara.ParagraphFormat.StyleName;
        targetPara.ParagraphFormat.Alignment = sourcePara.ParagraphFormat.Alignment;
        targetPara.ParagraphFormat.LeftIndent = sourcePara.ParagraphFormat.LeftIndent;
        targetPara.ParagraphFormat.RightIndent = sourcePara.ParagraphFormat.RightIndent;
        targetPara.ParagraphFormat.FirstLineIndent = sourcePara.ParagraphFormat.FirstLineIndent;
        targetPara.ParagraphFormat.SpaceBefore = sourcePara.ParagraphFormat.SpaceBefore;
        targetPara.ParagraphFormat.SpaceAfter = sourcePara.ParagraphFormat.SpaceAfter;
        targetPara.ParagraphFormat.LineSpacing = sourcePara.ParagraphFormat.LineSpacing;
        targetPara.ParagraphFormat.LineSpacingRule = sourcePara.ParagraphFormat.LineSpacingRule;

        targetPara.ParagraphFormat.TabStops.Clear();
        for (var i = 0; i < sourcePara.ParagraphFormat.TabStops.Count; i++)
        {
            var tabStop = sourcePara.ParagraphFormat.TabStops[i];
            targetPara.ParagraphFormat.TabStops.Add(tabStop.Position, tabStop.Alignment, tabStop.Leader);
        }

        doc.Save(outputPath);

        var result = "Paragraph format copied successfully\n";
        result += $"Source paragraph: #{sourceParagraphIndex}\n";
        result += $"Target paragraph: #{targetParagraphIndex}\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Merges multiple paragraphs into one
    /// </summary>
    /// <param name="arguments">JSON arguments containing startIndex, endIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> MergeParagraphsAsync(JsonObject? arguments, string path, string outputPath)
    {
        var startParagraphIndex = ArgumentHelper.GetInt(arguments, "startParagraphIndex");
        var endParagraphIndex = ArgumentHelper.GetInt(arguments, "endParagraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Start paragraph index {startParagraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"End paragraph index {endParagraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (startParagraphIndex > endParagraphIndex)
            throw new ArgumentException(
                $"Start paragraph index {startParagraphIndex} cannot be greater than end paragraph index {endParagraphIndex}");

        if (startParagraphIndex == endParagraphIndex)
            throw new ArgumentException("Start and end paragraph indices are the same, no merge needed");

        var startPara = paragraphs[startParagraphIndex] as Paragraph;
        if (startPara == null) throw new InvalidOperationException("Unable to get start paragraph");

        for (var i = startParagraphIndex + 1; i <= endParagraphIndex; i++)
            if (paragraphs[i] is Paragraph para)
            {
                if (startPara.Runs.Count > 0)
                {
                    var spaceRun = new Run(doc, " ");
                    startPara.AppendChild(spaceRun);
                }

                var runsToMove = para.Runs.ToArray();
                foreach (var run in runsToMove) startPara.AppendChild(run);

                para.Remove();
            }

        doc.Save(outputPath);

        var result = "Paragraphs merged successfully\n";
        result += $"Merge range: Paragraph #{startParagraphIndex} to #{endParagraphIndex}\n";
        result += $"Merged paragraphs: {endParagraphIndex - startParagraphIndex + 1}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" => ParagraphAlignment.Justify,
            _ => ParagraphAlignment.Left
        };
    }

    private LineSpacingRule GetLineSpacingRule(string rule)
    {
        return rule.ToLower() switch
        {
            "single" => LineSpacingRule.Exactly,
            "oneAndHalf" => LineSpacingRule.Exactly,
            "double" => LineSpacingRule.Exactly,
            "atLeast" => LineSpacingRule.AtLeast,
            "exactly" => LineSpacingRule.Exactly,
            "multiple" => LineSpacingRule.Multiple,
            _ => LineSpacingRule.Exactly
        };
    }

    private TabAlignment GetTabAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => TabAlignment.Left,
            "center" => TabAlignment.Center,
            "right" => TabAlignment.Right,
            "decimal" => TabAlignment.Decimal,
            "bar" => TabAlignment.Bar,
            "clear" => TabAlignment.Clear,
            _ => TabAlignment.Left
        };
    }

    private TabLeader GetTabLeader(string leader)
    {
        return leader.ToLower() switch
        {
            "none" => TabLeader.None,
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            "heavy" => TabLeader.Heavy,
            "middleDot" => TabLeader.MiddleDot,
            _ => TabLeader.None
        };
    }
}