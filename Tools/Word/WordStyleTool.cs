using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing styles in Word documents (get, create, apply, copy)
/// </summary>
public class WordStyleTool : IAsposeTool
{
    public string Description =>
        @"Manage styles in Word documents. Supports 4 operations: get_styles, create_style, apply_style, copy_styles.

Usage examples:
- Get styles: word_style(operation='get_styles', path='doc.docx', includeBuiltIn=true)
- Create style: word_style(operation='create_style', path='doc.docx', styleName='CustomStyle', styleType='paragraph', fontSize=14, bold=true)
- Apply style: word_style(operation='apply_style', path='doc.docx', styleName='Heading 1', paragraphIndex=0)
- Copy styles: word_style(operation='copy_styles', path='doc.docx', sourcePath='template.docx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get_styles': Get all styles (required params: path)
- 'create_style': Create a new style (required params: path, styleName, styleType)
- 'apply_style': Apply style to paragraph (required params: path, styleName, paragraphIndex)
- 'copy_styles': Copy styles from another document (required params: path, sourcePath)",
                @enum = new[] { "get_styles", "create_style", "apply_style", "copy_styles" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            includeBuiltIn = new
            {
                type = "boolean",
                description = "Include built-in styles (for get_styles, default: false)"
            },
            styleName = new
            {
                type = "string",
                description = "Style name (required for create_style, apply_style)"
            },
            styleType = new
            {
                type = "string",
                description = "Style type: paragraph, character, table, list (for create_style, default: paragraph)",
                @enum = new[] { "paragraph", "character", "table", "list" }
            },
            baseStyle = new
            {
                type = "string",
                description = "Base style to inherit from (for create_style)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (for create_style)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (for create_style)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (for create_style)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (for create_style)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (for create_style)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (for create_style)"
            },
            underline = new
            {
                type = "boolean",
                description = "Underline text (for create_style)"
            },
            color = new
            {
                type = "string",
                description = "Text color hex (for create_style)"
            },
            alignment = new
            {
                type = "string",
                description = "Paragraph alignment: left, center, right, justify (for create_style)",
                @enum = new[] { "left", "center", "right", "justify" }
            },
            spaceBefore = new
            {
                type = "number",
                description = "Space before paragraph in points (for create_style)"
            },
            spaceAfter = new
            {
                type = "number",
                description = "Space after paragraph in points (for create_style)"
            },
            lineSpacing = new
            {
                type = "number",
                description = "Line spacing multiplier (for create_style)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, for apply_style)"
            },
            paragraphIndices = new
            {
                type = "array",
                description = "Array of paragraph indices (for apply_style)",
                items = new { type = "number" }
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, for apply_style, default: 0)"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index (0-based, for apply_style)"
            },
            applyToAllParagraphs = new
            {
                type = "boolean",
                description = "Apply to all paragraphs (for apply_style, default: false)"
            },
            sourceDocument = new
            {
                type = "string",
                description = "Source document path to copy styles from (for copy_styles)"
            },
            styleNames = new
            {
                type = "array",
                description = "Array of style names to copy (for copy_styles, if not provided copies all)",
                items = new { type = "string" }
            },
            overwriteExisting = new
            {
                type = "boolean",
                description = "Overwrite existing styles (for copy_styles, default: false)"
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

        return operation.ToLower() switch
        {
            "get_styles" => await GetStyles(arguments),
            "create_style" => await CreateStyle(arguments),
            "apply_style" => await ApplyStyle(arguments),
            "copy_styles" => await CopyStyles(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets all styles from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional includeBuiltIn</param>
    /// <returns>Formatted string with all styles</returns>
    private Task<string> GetStyles(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            var includeBuiltIn = ArgumentHelper.GetBool(arguments, "includeBuiltIn", false);

            var doc = new Document(path);
            var result = new StringBuilder();

            result.AppendLine("=== Document Styles ===\n");

            List<Style> paraStyles;

            if (includeBuiltIn)
            {
                // Include all styles (built-in and custom)
                paraStyles = doc.Styles
                    .Where(s => s.Type == StyleType.Paragraph)
                    .OrderBy(s => s.Name)
                    .ToList();
            }
            else
            {
                var usedStyleNames = new HashSet<string>();
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>();
                foreach (var para in paragraphs)
                    if (para.ParagraphFormat.Style != null && !string.IsNullOrEmpty(para.ParagraphFormat.Style.Name))
                        usedStyleNames.Add(para.ParagraphFormat.Style.Name);

                // Return styles that are either custom (not built-in) OR are built-in but actually used in the document
                paraStyles = doc.Styles
                    .Where(s => s.Type == StyleType.Paragraph && (!s.BuiltIn || usedStyleNames.Contains(s.Name)))
                    .OrderBy(s => s.Name)
                    .ToList();
            }

            result.AppendLine("【Paragraph Styles】");
            if (paraStyles.Count == 0)
                result.AppendLine("(No paragraph styles found)");
            else
                foreach (var style in paraStyles)
                {
                    result.AppendLine($"\nStyle Name: {style.Name}");
                    result.AppendLine($"  Built-in: {(style.BuiltIn ? "Yes" : "No")}");
                    if (!string.IsNullOrEmpty(style.BaseStyleName))
                        result.AppendLine($"  Based on: {style.BaseStyleName}");

                    var font = style.Font;
                    if (font.NameAscii != font.NameFarEast)
                    {
                        result.AppendLine($"  Font (ASCII): {font.NameAscii}");
                        result.AppendLine($"  Font (Far East): {font.NameFarEast}");
                    }
                    else
                    {
                        result.AppendLine($"  Font: {font.Name}");
                    }

                    result.AppendLine($"  Size: {font.Size} pt");
                    if (font.Bold) result.AppendLine("  Bold: Yes");
                    if (font.Italic) result.AppendLine("  Italic: Yes");

                    var paraFormat = style.ParagraphFormat;
                    result.AppendLine($"  Alignment: {paraFormat.Alignment}");
                    if (paraFormat.SpaceBefore != 0)
                        result.AppendLine($"  Space Before: {paraFormat.SpaceBefore} pt");
                    if (paraFormat.SpaceAfter != 0)
                        result.AppendLine($"  Space After: {paraFormat.SpaceAfter} pt");
                }

            result.AppendLine($"\n\nTotal Paragraph Styles: {paraStyles.Count}");
            if (!includeBuiltIn)
                result.AppendLine("(Showing custom styles and built-in styles actually used in the document)");

            return result.ToString();
        });
    }

    /// <summary>
    ///     Creates a new style
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing path, styleName, styleType, optional baseStyleName, formatting
    ///     options, outputPath
    /// </param>
    /// <returns>Success message</returns>
    private Task<string> CreateStyle(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
            var styleName = ArgumentHelper.GetString(arguments, "styleName");
            var styleTypeStr = ArgumentHelper.GetString(arguments, "styleType", "paragraph");
            var baseStyle = ArgumentHelper.GetStringNullable(arguments, "baseStyle");
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
            var underline = ArgumentHelper.GetBoolNullable(arguments, "underline");
            var color = ArgumentHelper.GetStringNullable(arguments, "color");
            var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment");
            var spaceBefore = ArgumentHelper.GetDoubleNullable(arguments, "spaceBefore");
            var spaceAfter = ArgumentHelper.GetDoubleNullable(arguments, "spaceAfter");
            var lineSpacing = ArgumentHelper.GetDoubleNullable(arguments, "lineSpacing");

            var doc = new Document(path);

            if (doc.Styles[styleName] != null)
                throw new InvalidOperationException($"Style '{styleName}' already exists");

            var styleType = styleTypeStr.ToLower() switch
            {
                "character" => StyleType.Character,
                "table" => StyleType.Table,
                "list" => StyleType.List,
                _ => StyleType.Paragraph
            };

            var style = doc.Styles.Add(styleType, styleName);

            if (!string.IsNullOrEmpty(baseStyle))
            {
                var baseStyleObj = doc.Styles[baseStyle];
                if (baseStyleObj != null)
                    style.BaseStyleName = baseStyle;
            }

            if (!string.IsNullOrEmpty(fontNameAscii))
                style.Font.NameAscii = fontNameAscii;

            if (!string.IsNullOrEmpty(fontNameFarEast))
                style.Font.NameFarEast = fontNameFarEast;

            if (!string.IsNullOrEmpty(fontName))
            {
                if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                {
                    style.Font.Name = fontName;
                }
                else
                {
                    if (string.IsNullOrEmpty(fontNameAscii))
                        style.Font.NameAscii = fontName;
                    if (string.IsNullOrEmpty(fontNameFarEast))
                        style.Font.NameFarEast = fontName;
                }
            }

            if (fontSize.HasValue)
                style.Font.Size = fontSize.Value;

            if (bold.HasValue)
                style.Font.Bold = bold.Value;

            if (italic.HasValue)
                style.Font.Italic = italic.Value;

            if (underline.HasValue)
                style.Font.Underline = underline.Value ? Underline.Single : Underline.None;

            if (!string.IsNullOrEmpty(color))
                // Parse color with error handling - throws ArgumentException on failure
                style.Font.Color = ColorHelper.ParseColor(color, true);

            if (styleType == StyleType.Paragraph || styleType == StyleType.List)
            {
                if (!string.IsNullOrEmpty(alignment))
                    style.ParagraphFormat.Alignment = alignment.ToLower() switch
                    {
                        "center" => ParagraphAlignment.Center,
                        "right" => ParagraphAlignment.Right,
                        "justify" => ParagraphAlignment.Justify,
                        _ => ParagraphAlignment.Left
                    };

                if (spaceBefore.HasValue)
                    style.ParagraphFormat.SpaceBefore = spaceBefore.Value;

                if (spaceAfter.HasValue)
                    style.ParagraphFormat.SpaceAfter = spaceAfter.Value;

                if (lineSpacing.HasValue)
                {
                    style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                    style.ParagraphFormat.LineSpacing = lineSpacing.Value * 12;
                }
            }

            doc.Save(outputPath);
            return $"Style '{styleName}' created successfully: {outputPath}";
        });
    }

    /// <summary>
    ///     Applies a style to paragraphs or runs
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, styleName, optional paragraphIndex, runIndex, outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> ApplyStyle(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
            var styleName = ArgumentHelper.GetString(arguments, "styleName");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
            var paragraphIndicesArray = ArgumentHelper.GetArray(arguments, "paragraphIndices", false);
            var tableIndex = ArgumentHelper.GetIntNullable(arguments, "tableIndex");
            var applyToAllParagraphs = ArgumentHelper.GetBool(arguments, "applyToAllParagraphs", false);

            var doc = new Document(path);
            var style = doc.Styles[styleName];
            if (style == null)
                throw new ArgumentException($"Style '{styleName}' not found");

            var appliedCount = 0;

            if (tableIndex.HasValue)
            {
                var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
                if (tableIndex.Value < 0 || tableIndex.Value >= tables.Count)
                    throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");
                tables[tableIndex.Value].Style = style;
                appliedCount = 1;
            }
            else if (applyToAllParagraphs)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                foreach (var para in paragraphs)
                {
                    var paraFormat = para.ParagraphFormat;
                    var isEmpty = string.IsNullOrWhiteSpace(para.GetText());

                    if (isEmpty)
                    {
                        paraFormat.ClearFormatting();
                        try
                        {
                            if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                                paraFormat.StyleIdentifier = style.StyleIdentifier;
                        }
                        catch (Exception ex)
                        {
                            // Ignore StyleIdentifier errors
                            Console.Error.WriteLine($"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
                        }
                    }

                    paraFormat.Style = style;
                    paraFormat.StyleName = styleName;

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
                            Console.Error.WriteLine($"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
                        }
                    }

                    appliedCount++;
                }
            }
            else if (paragraphIndicesArray is { Count: > 0 })
            {
                var sectionIdx = sectionIndex ?? 0;
                if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
                    throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

                var section = doc.Sections[sectionIdx];
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

                foreach (var idxObj in paragraphIndicesArray)
                {
                    var idx = idxObj?.GetValue<int>();
                    if (idx is >= 0 && idx.Value < paragraphs.Count)
                    {
                        var para = paragraphs[idx.Value];
                        var paraFormat = para.ParagraphFormat;
                        var isEmpty = string.IsNullOrWhiteSpace(para.GetText());

                        if (isEmpty)
                        {
                            paraFormat.ClearFormatting();
                            try
                            {
                                if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                                    paraFormat.StyleIdentifier = style.StyleIdentifier;
                            }
                            catch (Exception ex)
                            {
                                // Ignore StyleIdentifier errors
                                Console.Error.WriteLine($"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
                            }
                        }

                        paraFormat.Style = style;
                        paraFormat.StyleName = styleName;

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
                                Console.Error.WriteLine($"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
                            }
                        }

                        appliedCount++;
                    }
                }
            }
            else if (paragraphIndex.HasValue)
            {
                var sectionIdx = sectionIndex ?? 0;
                if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
                    throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

                var section = doc.Sections[sectionIdx];
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

                if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                    throw new ArgumentException(
                        $"paragraphIndex must be between 0 and {paragraphs.Count - 1} (section {sectionIdx} has {paragraphs.Count} paragraphs, total document paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count})");

                var para = paragraphs[paragraphIndex.Value];
                var paraFormat = para.ParagraphFormat;

                // For empty paragraphs, we need to ensure the style is properly applied
                // Use StyleIdentifier for more reliable style application
                var isEmpty = string.IsNullOrWhiteSpace(para.GetText());

                if (isEmpty)
                {
                    // For empty paragraphs, clear formatting first, then apply style using StyleIdentifier
                    paraFormat.ClearFormatting();

                    // Try to use StyleIdentifier if the style has one
                    try
                    {
                        if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                            paraFormat.StyleIdentifier = style.StyleIdentifier;
                    }
                    catch
                    {
                        // If StyleIdentifier fails, continue with Style and StyleName
                    }
                }

                // Apply style directly to paragraph format
                paraFormat.Style = style;
                // Also set StyleName to ensure it's properly applied (especially for empty paragraphs)
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
                    catch
                    {
                        // Ignore StyleIdentifier errors
                    }
                }

                appliedCount = 1;
            }
            else
            {
                throw new ArgumentException(
                    "Either paragraphIndex, paragraphIndices, tableIndex, or applyToAllParagraphs must be provided");
            }

            doc.Save(outputPath);
            return $"Applied style '{styleName}' to {appliedCount} element(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Copies styles from source document to destination document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, sourcePath, optional outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> CopyStyles(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
            var sourceDocument = ArgumentHelper.GetString(arguments, "sourceDocument");
            SecurityHelper.ValidateFilePath(sourceDocument, "sourceDocument", true);
            var overwriteExisting = ArgumentHelper.GetBool(arguments, "overwriteExisting", false);

            if (!File.Exists(sourceDocument))
                throw new FileNotFoundException($"Source document not found: {sourceDocument}");

            var targetDoc = new Document(path);
            var sourceDoc = new Document(sourceDocument);

            var styleNames = new List<string>();
            if (arguments?.ContainsKey("styleNames") == true)
            {
                var stylesArray = arguments["styleNames"]?.AsArray();
                if (stylesArray != null)
                    foreach (var item in stylesArray)
                    {
                        var name = item?.GetValue<string>();
                        if (!string.IsNullOrEmpty(name))
                            styleNames.Add(name);
                    }
            }

            var copyAll = styleNames.Count == 0;
            var copiedCount = 0;
            var skippedCount = 0;

            foreach (var sourceStyle in sourceDoc.Styles)
            {
                if (!copyAll && !styleNames.Contains(sourceStyle.Name))
                    continue;

                var existingStyle = targetDoc.Styles[sourceStyle.Name];

                if (existingStyle != null && !overwriteExisting)
                {
                    skippedCount++;
                    continue;
                }

                try
                {
                    if (existingStyle != null && overwriteExisting)
                    {
                        CopyStyleProperties(sourceStyle, existingStyle);
                    }
                    else
                    {
                        var newStyle = targetDoc.Styles.Add(sourceStyle.Type, sourceStyle.Name);
                        CopyStyleProperties(sourceStyle, newStyle);
                    }

                    copiedCount++;
                }
                catch (Exception ex)
                {
                    skippedCount++;
                    Console.Error.WriteLine($"[WARN] Failed to copy style '{sourceStyle.Name}': {ex.Message}");
                }
            }

            targetDoc.Save(outputPath);
            return
                $"Copied {copiedCount} style(s) from {Path.GetFileName(sourceDocument)}. Skipped: {skippedCount}. Output: {outputPath}";
        });
    }

    private void CopyStyleProperties(Style sourceStyle, Style targetStyle)
    {
        targetStyle.Font.Name = sourceStyle.Font.Name;
        targetStyle.Font.NameAscii = sourceStyle.Font.NameAscii;
        targetStyle.Font.NameFarEast = sourceStyle.Font.NameFarEast;
        targetStyle.Font.Size = sourceStyle.Font.Size;
        targetStyle.Font.Bold = sourceStyle.Font.Bold;
        targetStyle.Font.Italic = sourceStyle.Font.Italic;
        targetStyle.Font.Color = sourceStyle.Font.Color;
        targetStyle.Font.Underline = sourceStyle.Font.Underline;

        if (sourceStyle.Type == StyleType.Paragraph)
        {
            targetStyle.ParagraphFormat.Alignment = sourceStyle.ParagraphFormat.Alignment;
            targetStyle.ParagraphFormat.SpaceBefore = sourceStyle.ParagraphFormat.SpaceBefore;
            targetStyle.ParagraphFormat.SpaceAfter = sourceStyle.ParagraphFormat.SpaceAfter;
            targetStyle.ParagraphFormat.LineSpacing = sourceStyle.ParagraphFormat.LineSpacing;
            targetStyle.ParagraphFormat.LineSpacingRule = sourceStyle.ParagraphFormat.LineSpacingRule;
            targetStyle.ParagraphFormat.LeftIndent = sourceStyle.ParagraphFormat.LeftIndent;
            targetStyle.ParagraphFormat.RightIndent = sourceStyle.ParagraphFormat.RightIndent;
            targetStyle.ParagraphFormat.FirstLineIndent = sourceStyle.ParagraphFormat.FirstLineIndent;
        }
    }
}