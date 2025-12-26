using System.Drawing;
using System.Text.Json;
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
- Copy format: word_paragraph(operation='copy_format', path='doc.docx', sourceParagraphIndex=0, targetParagraphIndex=1)
- Merge paragraphs: word_paragraph(operation='merge', path='doc.docx', startParagraphIndex=0, endParagraphIndex=2)

Important notes for 'get' operation:
- By default, returns ALL paragraphs in the document structure, including paragraphs inside Comment objects, table cells, and TextBoxes
- Use includeCommentParagraphs=false to get only Body paragraphs (visible in document body, excluding table cells and comments)
- Each paragraph shows its ParentNode type to help identify its location
- Paragraphs inside Comment objects are marked with '[Comment]' in the location field
- Paragraphs inside table cells are marked with '[Cell]' in the location field
- Paragraphs inside TextBoxes are marked with '[TextBox]' in the location field
- To check paragraph styles in table cells, use includeCommentParagraphs=true";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'insert': Insert a new paragraph (required params: path, text; optional: paragraphIndex)
- 'delete': Delete a paragraph (required params: path, paragraphIndex)
- 'edit': Edit paragraph format (required params: path, paragraphIndex)
- 'get': Get paragraph content (required params: path; optional: paragraphIndex)
- 'get_format': Get paragraph format (required params: path, paragraphIndex)
- 'copy_format': Copy paragraph format (required params: path, sourceParagraphIndex, targetParagraphIndex)
- 'merge': Merge paragraphs (required params: path, startParagraphIndex, endParagraphIndex)",
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
                    @"Include paragraphs inside nested structures (optional, default: true, for get operation).
When true: Returns ALL paragraphs in the document structure, including:
  - Paragraphs inside Comment objects (marked with [Comment])
  - Paragraphs inside table cells (marked with [Cell])
  - Paragraphs inside TextBox/Shape objects (if includeTextboxParagraphs is also true)
When false: Returns only Body paragraphs (direct children of document Body, visible in document body).
Note: To check paragraph styles in table cells (e.g., after using add_table with cellStyles), 
      you must set includeCommentParagraphs=true. This returns the document's underlying structure, 
      including nested content that is not direct children of Body."
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
                    "Line spacing value. For 'multiple/single/oneAndHalf/double' rules, use multiplier (e.g., 1.0, 1.5, 2.0). For 'atLeast/exactly' rules, use points. Optional, for edit operation."
            },
            lineSpacingRule = new
            {
                type = "string",
                description =
                    "Line spacing rule: 'single' (1x), 'oneAndHalf' (1.5x), 'double' (2x), 'atLeast' (minimum points), 'exactly' (fixed points), 'multiple' (custom multiplier). Optional, for edit operation.",
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
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation switch
        {
            "insert" => await InsertParagraphAsync(path, outputPath, arguments),
            "delete" => await DeleteParagraphAsync(path, outputPath, arguments),
            "edit" => await EditParagraphAsync(path, outputPath, arguments),
            "get" => await GetParagraphsAsync(path, arguments),
            "get_format" => await GetParagraphFormatAsync(path, arguments),
            "copy_format" => await CopyParagraphFormatAsync(path, outputPath, arguments),
            "merge" => await MergeParagraphsAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a paragraph into the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, optional paragraphIndex, styleName, formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> InsertParagraphAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
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
                    var validRange = paragraphs.Count > 0 ? $"0-{paragraphs.Count - 1}" : "none (document is empty)";
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: {validRange}, or -1 for beginning).");
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

            // Use previously cached paragraphs.Count + 1 to avoid redundant GetChildNodes call
            result += $"Document paragraph count: {paragraphs.Count + 1}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes a paragraph from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex (0-based, or -1 for last paragraph)</param>
    /// <returns>Success message with deleted paragraph preview</returns>
    private Task<string> DeleteParagraphAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
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
                    $"Paragraph index {paragraphIndex} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}, or -1 for last).");

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

            return result;
        });
    }

    /// <summary>
    ///     Edits paragraph properties
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional text, formatting options, sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> EditParagraphAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
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
                else
                    // For Multiple rule, LineSpacing is a multiplier (1.0, 1.5, 2.0)
                    // For AtLeast/Exactly, LineSpacing is in points
                    paraFormat.LineSpacing = lineSpacingRule.ToLower() switch
                    {
                        "single" => 1.0,
                        "oneandhalf" => 1.5,
                        "double" => 2.0,
                        _ => 1.0
                    };
            }

            var styleName = ArgumentHelper.GetStringNullable(arguments, "styleName");
            if (!string.IsNullOrEmpty(styleName))
                try
                {
                    var style = doc.Styles[styleName];
                    if (style != null)
                    {
                        // For empty paragraphs, we need to ensure the style is properly applied
                        // Use StyleIdentifier for more reliable style application
                        var isEmpty = string.IsNullOrWhiteSpace(para.GetText());

                        if (isEmpty)
                        {
                            // For empty paragraphs, clear formatting first, then apply style using StyleIdentifier
                            paraFormat.ClearFormatting();

                            try
                            {
                                if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                                    paraFormat.StyleIdentifier = style.StyleIdentifier;
                            }
                            catch (Exception ex)
                            {
                                // If StyleIdentifier fails, continue with Style and StyleName
                                Console.Error.WriteLine(
                                    $"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
                            }
                        }

                        // Apply style directly to paragraph format
                        paraFormat.Style = style;
                        // Also set StyleName to ensure it's properly applied (especially for paragraphs in table cells and empty paragraphs)
                        paraFormat.StyleName = styleName;

                        // For empty paragraphs, force style application by clearing and re-applying
                        if (isEmpty)
                        {
                            paraFormat.ClearFormatting();
                            paraFormat.Style = style;
                            paraFormat.StyleName = styleName;
                            try
                            {
                                if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                                    paraFormat.StyleIdentifier = style.StyleIdentifier;
                            }
                            catch (Exception ex)
                            {
                                // Ignore StyleIdentifier errors
                                Console.Error.WriteLine(
                                    $"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        throw new ArgumentException(
                            $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
                    }
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
                    // For empty paragraphs, check if user specified any font settings
                    var hasFontSettings = fontName != null || fontNameAscii != null || fontNameFarEast != null ||
                                          fontSize.HasValue || bold.HasValue || italic.HasValue ||
                                          underlineStr != null || colorStr != null;

                    if (hasFontSettings)
                    {
                        // Create a sentinel Run to preserve font settings for future user input in Word
                        // Using zero-width space (U+200B) so it's invisible but preserves formatting
                        var sentinelRun = new Run(doc, "\u200B");
                        FontHelper.Word.ApplyFontSettings(
                            sentinelRun,
                            fontName,
                            fontNameAscii,
                            fontNameFarEast,
                            fontSize,
                            bold,
                            italic,
                            underlineStr,
                            colorStr
                        );
                        para.AppendChild(sentinelRun);
                    }
                }
                else
                {
                    foreach (var run in runs)
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
            return resultMsg;
        });
    }

    /// <summary>
    ///     Gets all paragraphs from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">
    ///     JSON arguments containing optional: sectionIndex, includeEmpty, styleFilter,
    ///     includeCommentParagraphs, includeTextboxParagraphs
    /// </param>
    /// <returns>JSON formatted string with all paragraphs including their location, style, and text preview</returns>
    private Task<string> GetParagraphsAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
            var includeEmpty = ArgumentHelper.GetBool(arguments, "includeEmpty", true);
            var styleFilter = ArgumentHelper.GetStringNullable(arguments, "styleFilter");
            var includeCommentParagraphs = ArgumentHelper.GetBool(arguments, "includeCommentParagraphs", true);
            var includeTextboxParagraphs = ArgumentHelper.GetBool(arguments, "includeTextboxParagraphs", true);

            var doc = new Document(path);

            List<Paragraph> paragraphs;
            if (sectionIndex.HasValue)
            {
                if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                    throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
                paragraphs = doc.Sections[sectionIndex.Value].Body
                    .GetChildNodes(NodeType.Paragraph, includeCommentParagraphs).Cast<Paragraph>().ToList();
            }
            else
            {
                if (includeCommentParagraphs)
                {
                    paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                }
                else
                {
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

            if (!includeTextboxParagraphs)
                paragraphs = paragraphs.Where(p =>
                {
                    var shapeAncestor = p.GetAncestor(NodeType.Shape);
                    if (shapeAncestor is Shape { ShapeType: ShapeType.TextBox })
                        return false;

                    var currentNode = p.ParentNode;
                    while (currentNode != null)
                    {
                        if (currentNode.NodeType == NodeType.Shape)
                            if (currentNode is Shape { ShapeType: ShapeType.TextBox })
                                return false;
                        currentNode = currentNode.ParentNode;
                    }

                    return true;
                }).ToList();

            var paragraphList = new List<object>();
            for (var i = 0; i < paragraphs.Count; i++)
            {
                var para = paragraphs[i];
                var text = para.GetText().Trim();
                var location = "Body";
                string? commentInfo = null;

                if (para.ParentNode != null)
                {
                    var commentAncestor = para.GetAncestor(NodeType.Comment);
                    if (commentAncestor != null)
                    {
                        location = "Comment";
                        if (commentAncestor is Comment comment)
                            commentInfo = $"ID: {comment.Id}, Author: {comment.Author}";
                    }
                    else
                    {
                        var shapeAncestor = para.GetAncestor(NodeType.Shape);
                        if (shapeAncestor != null)
                        {
                            location = shapeAncestor is Shape { ShapeType: ShapeType.TextBox } ? "TextBox" : "Shape";
                        }
                        else
                        {
                            var bodyAncestor = para.GetAncestor(NodeType.Body);
                            if (bodyAncestor == null || para.ParentNode.NodeType != NodeType.Body)
                                location = para.ParentNode.NodeType.ToString();
                        }
                    }
                }

                var paraInfo = new Dictionary<string, object?>
                {
                    ["index"] = i,
                    ["location"] = location,
                    ["style"] = para.ParagraphFormat.Style?.Name,
                    ["text"] = text.Length > 100 ? text[..100] + "..." : text,
                    ["textLength"] = text.Length
                };

                if (commentInfo != null) paraInfo["commentInfo"] = commentInfo;

                paragraphList.Add(paraInfo);
            }

            var result = new
            {
                count = paragraphs.Count,
                filters = new
                {
                    sectionIndex,
                    includeEmpty,
                    styleFilter,
                    includeCommentParagraphs,
                    includeTextboxParagraphs
                },
                paragraphs = paragraphList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Gets format information for a paragraph
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex (0-based), optional includeRunDetails</param>
    /// <returns>JSON formatted string with paragraph format details including alignment, indentation, spacing, and font info</returns>
    private Task<string> GetParagraphFormatAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var includeRunDetails = ArgumentHelper.GetBool(arguments, "includeRunDetails", true);

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

            var para = paragraphs[paragraphIndex] as Paragraph;
            if (para == null)
                throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex}");

            var format = para.ParagraphFormat;
            var text = para.GetText().Trim();

            // Build result object
            var resultDict = new Dictionary<string, object?>
            {
                ["paragraphIndex"] = paragraphIndex,
                ["text"] = text,
                ["textLength"] = text.Length,
                ["runCount"] = para.Runs.Count,
                ["paragraphFormat"] = new
                {
                    styleName = format.StyleName,
                    alignment = format.Alignment.ToString(),
                    leftIndent = Math.Round(format.LeftIndent, 2),
                    rightIndent = Math.Round(format.RightIndent, 2),
                    firstLineIndent = Math.Round(format.FirstLineIndent, 2),
                    spaceBefore = Math.Round(format.SpaceBefore, 2),
                    spaceAfter = Math.Round(format.SpaceAfter, 2),
                    lineSpacing = Math.Round(format.LineSpacing, 2),
                    lineSpacingRule = format.LineSpacingRule.ToString()
                }
            };

            // List format
            if (para.ListFormat is { IsListItem: true })
                resultDict["listFormat"] = new
                {
                    isListItem = true,
                    listLevel = para.ListFormat.ListLevelNumber,
                    listId = para.ListFormat.List?.ListId
                };

            // Borders
            var borders = new Dictionary<string, object>();
            if (format.Borders.Top.LineStyle != LineStyle.None)
                borders["top"] = new
                {
                    lineStyle = format.Borders.Top.LineStyle.ToString(), lineWidth = format.Borders.Top.LineWidth,
                    color = format.Borders.Top.Color.Name
                };
            if (format.Borders.Bottom.LineStyle != LineStyle.None)
                borders["bottom"] = new
                {
                    lineStyle = format.Borders.Bottom.LineStyle.ToString(), lineWidth = format.Borders.Bottom.LineWidth,
                    color = format.Borders.Bottom.Color.Name
                };
            if (format.Borders.Left.LineStyle != LineStyle.None)
                borders["left"] = new
                {
                    lineStyle = format.Borders.Left.LineStyle.ToString(), lineWidth = format.Borders.Left.LineWidth,
                    color = format.Borders.Left.Color.Name
                };
            if (format.Borders.Right.LineStyle != LineStyle.None)
                borders["right"] = new
                {
                    lineStyle = format.Borders.Right.LineStyle.ToString(), lineWidth = format.Borders.Right.LineWidth,
                    color = format.Borders.Right.Color.Name
                };
            if (borders.Count > 0)
                resultDict["borders"] = borders;

            // Background color
            if (format.Shading.BackgroundPatternColor.ToArgb() != Color.Empty.ToArgb())
            {
                var color = format.Shading.BackgroundPatternColor;
                resultDict["backgroundColor"] = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }

            // Tab stops
            if (format.TabStops.Count > 0)
            {
                var tabStops = new List<object>();
                for (var i = 0; i < format.TabStops.Count; i++)
                {
                    var tab = format.TabStops[i];
                    tabStops.Add(new
                    {
                        position = Math.Round(tab.Position, 2), alignment = tab.Alignment.ToString(),
                        leader = tab.Leader.ToString()
                    });
                }

                resultDict["tabStops"] = tabStops;
            }

            // First run font format
            if (para.Runs.Count > 0)
            {
                var firstRun = para.Runs[0];
                var fontInfo = new Dictionary<string, object?>
                {
                    ["fontSize"] = firstRun.Font.Size
                };

                if (firstRun.Font.NameAscii != firstRun.Font.NameFarEast)
                {
                    fontInfo["fontAscii"] = firstRun.Font.NameAscii;
                    fontInfo["fontFarEast"] = firstRun.Font.NameFarEast;
                }
                else
                {
                    fontInfo["font"] = firstRun.Font.Name;
                }

                if (firstRun.Font.Bold) fontInfo["bold"] = true;
                if (firstRun.Font.Italic) fontInfo["italic"] = true;
                if (firstRun.Font.Underline != Underline.None)
                    fontInfo["underline"] = firstRun.Font.Underline.ToString();
                if (firstRun.Font.StrikeThrough) fontInfo["strikethrough"] = true;
                if (firstRun.Font.Superscript) fontInfo["superscript"] = true;
                if (firstRun.Font.Subscript) fontInfo["subscript"] = true;
                if (firstRun.Font.Color.ToArgb() != Color.Empty.ToArgb())
                    fontInfo["color"] =
                        $"#{firstRun.Font.Color.R:X2}{firstRun.Font.Color.G:X2}{firstRun.Font.Color.B:X2}";
                if (firstRun.Font.HighlightColor != Color.Empty)
                    fontInfo["highlightColor"] = firstRun.Font.HighlightColor.Name;

                resultDict["fontFormat"] = fontInfo;
            }

            // Run details
            if (includeRunDetails && para.Runs.Count > 1)
            {
                var runs = new List<object>();
                for (var i = 0; i < Math.Min(para.Runs.Count, 10); i++)
                {
                    var run = para.Runs[i];
                    var runInfo = new Dictionary<string, object?>
                    {
                        ["index"] = i,
                        ["text"] = run.Text.Replace("\r", "\\r").Replace("\n", "\\n"),
                        ["fontSize"] = run.Font.Size
                    };

                    if (run.Font.NameAscii != run.Font.NameFarEast)
                    {
                        runInfo["fontAscii"] = run.Font.NameAscii;
                        runInfo["fontFarEast"] = run.Font.NameFarEast;
                    }
                    else
                    {
                        runInfo["font"] = run.Font.Name;
                    }

                    if (run.Font.Bold) runInfo["bold"] = true;
                    if (run.Font.Italic) runInfo["italic"] = true;
                    if (run.Font.Underline != Underline.None) runInfo["underline"] = run.Font.Underline.ToString();

                    runs.Add(runInfo);
                }

                resultDict["runs"] = new
                {
                    total = para.Runs.Count,
                    displayed = Math.Min(para.Runs.Count, 10),
                    details = runs
                };
            }

            return JsonSerializer.Serialize(resultDict, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Copies paragraph format from source to target paragraph
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sourceParagraphIndex and targetParagraphIndex (0-based)</param>
    /// <returns>Success message with source and target paragraph indices</returns>
    private Task<string> CopyParagraphFormatAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sourceParagraphIndex = ArgumentHelper.GetInt(arguments, "sourceParagraphIndex");
            var targetParagraphIndex = ArgumentHelper.GetInt(arguments, "targetParagraphIndex");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (sourceParagraphIndex < 0 || sourceParagraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Source paragraph index {sourceParagraphIndex} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

            if (targetParagraphIndex < 0 || targetParagraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Target paragraph index {targetParagraphIndex} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

            var sourcePara = paragraphs[sourceParagraphIndex] as Paragraph;
            var targetPara = paragraphs[targetParagraphIndex] as Paragraph;

            if (sourcePara == null || targetPara == null)
                throw new InvalidOperationException("Unable to get paragraphs");

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

            return result;
        });
    }

    /// <summary>
    ///     Merges multiple paragraphs into one by combining their content
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing startParagraphIndex and endParagraphIndex (0-based, inclusive)</param>
    /// <returns>Success message with merge range and remaining paragraph count</returns>
    private Task<string> MergeParagraphsAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startParagraphIndex = ArgumentHelper.GetInt(arguments, "startParagraphIndex");
            var endParagraphIndex = ArgumentHelper.GetInt(arguments, "endParagraphIndex");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Start paragraph index {startParagraphIndex} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

            if (endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"End paragraph index {endParagraphIndex} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

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

            return result;
        });
    }

    /// <summary>
    ///     Converts an alignment string to ParagraphAlignment enum
    /// </summary>
    /// <param name="alignment">Alignment string: left, center, right, or justify</param>
    /// <returns>Corresponding ParagraphAlignment enum value, defaults to Left</returns>
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

    /// <summary>
    ///     Converts a line spacing rule string to LineSpacingRule enum.
    ///     Note: For "single", "oneAndHalf", and "double", we use Multiple rule
    ///     with appropriate LineSpacing values (1.0, 1.5, 2.0 respectively as multipliers).
    ///     Using Exactly for these would cause fixed line height issues when font size changes.
    /// </summary>
    /// <param name="rule">Line spacing rule string: single, oneandhalf, double, atleast, exactly, or multiple</param>
    /// <returns>Corresponding LineSpacingRule enum value, defaults to Multiple</returns>
    private LineSpacingRule GetLineSpacingRule(string rule)
    {
        return rule.ToLower() switch
        {
            "atleast" => LineSpacingRule.AtLeast,
            "exactly" => LineSpacingRule.Exactly,
            _ => LineSpacingRule.Multiple
        };
    }

    /// <summary>
    ///     Converts a tab alignment string to TabAlignment enum
    /// </summary>
    /// <param name="alignment">Tab alignment string: left, center, right, decimal, bar, or clear</param>
    /// <returns>Corresponding TabAlignment enum value, defaults to Left</returns>
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

    /// <summary>
    ///     Converts a tab leader string to TabLeader enum
    /// </summary>
    /// <param name="leader">Tab leader string: none, dots, dashes, line, heavy, or middledot</param>
    /// <returns>Corresponding TabLeader enum value, defaults to None</returns>
    private TabLeader GetTabLeader(string leader)
    {
        return leader.ToLower() switch
        {
            "none" => TabLeader.None,
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            "heavy" => TabLeader.Heavy,
            "middledot" => TabLeader.MiddleDot,
            _ => TabLeader.None
        };
    }
}