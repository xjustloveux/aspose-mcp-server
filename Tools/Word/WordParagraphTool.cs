using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for paragraph operations in Word documents
/// Merges: WordInsertParagraphTool, WordDeleteParagraphTool, WordEditParagraphTool,
/// WordGetParagraphsTool, WordGetParagraphFormatTool, WordCopyParagraphFormatTool, WordMergeParagraphsTool
/// </summary>
public class WordParagraphTool : IAsposeTool
{
    public string Description => @"Manage paragraphs in Word documents. Supports 7 operations: insert, delete, edit, get, get_format, copy_format, merge.

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
                description = "Paragraph index (0-based, required for delete, edit, get_format operations, optional for insert/get operations). Valid range: 0 to (total paragraphs - 1), or -1 for last paragraph. Note: After delete operations, subsequent paragraph indices will shift automatically."
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
                description = "Style name to apply (e.g., 'Heading 1', '標題1', 'Normal', optional, for insert/edit operations)"
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
                description = "Include paragraphs inside Comment objects (optional, default: true, for get operation). Set to false to get only Body paragraphs (visible in document body). Note: This returns the document's underlying structure, including Comment content that is not visible in the body."
            },
            includeTextboxParagraphs = new
            {
                type = "boolean",
                description = "Include paragraphs inside TextBox/Shape objects (optional, default: true, for get operation). Set to false to exclude textbox paragraphs. Note: Textbox paragraphs are marked with [TextBox] in the location field."
            },
            // Get format parameters
            includeRunDetails = new
            {
                type = "boolean",
                description = "Include detailed run-level formatting (optional, default: true, for get_format operation)"
            },
            // Edit parameters
            fontName = new
            {
                type = "string",
                description = "Font name (e.g., '標楷體', 'Arial', optional, for edit operation)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, optional, for edit operation)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, optional, for edit operation)"
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
                description = "First line indent in points (positive for indent, negative for hanging, optional, for insert/edit operations)"
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
                description = "Line spacing (points or multiplier depending on lineSpacingRule, optional, for edit operation)"
            },
            lineSpacingRule = new
            {
                type = "string",
                description = "Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple (optional, for edit operation)",
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
                        alignment = new { type = "string", @enum = new[] { "left", "center", "right", "decimal", "bar", "clear" } },
                        leader = new { type = "string", @enum = new[] { "none", "dots", "dashes", "line", "heavy", "middleDot" } }
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
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
    /// Inserts a paragraph into the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing text, optional paragraphIndex, styleName, formatting options</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> InsertParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = ArgumentHelper.GetString(arguments, "text", "text");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var styleName = arguments?["styleName"]?.GetValue<string>();
        var alignment = arguments?["alignment"]?.GetValue<string>();
        var indentLeft = arguments?["indentLeft"]?.GetValue<double>();
        var indentRight = arguments?["indentRight"]?.GetValue<double>();
        var firstLineIndent = arguments?["firstLineIndent"]?.GetValue<double>();
        var spaceBefore = arguments?["spaceBefore"]?.GetValue<double>();
        var spaceAfter = arguments?["spaceAfter"]?.GetValue<double>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        Paragraph? targetPara = null;
        string insertPosition = "文檔末尾";
        
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[0] as Paragraph;
                    insertPosition = "文檔開頭";
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                insertPosition = $"段落 #{paragraphIndex.Value} 之後";
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
        }

        var para = new Paragraph(doc);
        var run = new Run(doc, text);
        para.AppendChild(run);

        if (!string.IsNullOrEmpty(styleName))
        {
            try
            {
                var style = doc.Styles[styleName];
                if (style != null)
                {
                    para.ParagraphFormat.StyleName = styleName;
                }
                else
                {
                    throw new ArgumentException($"找不到樣式 '{styleName}'，可用樣式請使用 word_get_styles 工具查看");
                }
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法應用樣式 '{styleName}': {ex.Message}，可用樣式請使用 word_get_styles 工具查看", ex);
            }
        }

        if (!string.IsNullOrEmpty(alignment))
        {
            para.ParagraphFormat.Alignment = alignment.ToLower() switch
            {
                "left" => ParagraphAlignment.Left,
                "right" => ParagraphAlignment.Right,
                "center" => ParagraphAlignment.Center,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };
        }
        
        // Apply indentation and spacing
        if (indentLeft.HasValue)
        {
            para.ParagraphFormat.LeftIndent = indentLeft.Value;
        }
        
        if (indentRight.HasValue)
        {
            para.ParagraphFormat.RightIndent = indentRight.Value;
        }
        
        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
        }
        
        if (spaceBefore.HasValue)
        {
            para.ParagraphFormat.SpaceBefore = spaceBefore.Value;
        }
        
        if (spaceAfter.HasValue)
        {
            para.ParagraphFormat.SpaceAfter = spaceAfter.Value;
        }

        if (targetPara != null)
        {
            if (paragraphIndex!.Value == -1)
            {
                targetPara.ParentNode.InsertBefore(para, targetPara);
            }
            else
            {
                targetPara.ParentNode.InsertAfter(para, targetPara);
            }
        }
        else
        {
            var body = doc.FirstSection.Body;
            body.AppendChild(para);
        }

        doc.Save(outputPath);

        var result = $"成功插入段落\n";
        result += $"插入位置: {insertPosition}\n";
        if (!string.IsNullOrEmpty(styleName))
        {
            result += $"應用樣式: {styleName}\n";
        }
        if (!string.IsNullOrEmpty(alignment))
        {
            result += $"對齊方式: {alignment}\n";
        }
        if (indentLeft.HasValue || indentRight.HasValue || firstLineIndent.HasValue)
        {
            result += $"縮排: ";
            var indentParts = new List<string>();
            if (indentLeft.HasValue) indentParts.Add($"左={indentLeft.Value}pt");
            if (indentRight.HasValue) indentParts.Add($"右={indentRight.Value}pt");
            if (firstLineIndent.HasValue) indentParts.Add($"首行={firstLineIndent.Value}pt");
            result += string.Join(", ", indentParts) + "\n";
        }
        if (spaceBefore.HasValue || spaceAfter.HasValue)
        {
            result += $"間距: ";
            var spaceParts = new List<string>();
            if (spaceBefore.HasValue) spaceParts.Add($"段前={spaceBefore.Value}pt");
            if (spaceAfter.HasValue) spaceParts.Add($"段後={spaceAfter.Value}pt");
            result += string.Join(", ", spaceParts) + "\n";
        }
        result += $"文檔段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    /// Deletes a paragraph from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", "paragraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        // Handle paragraphIndex=-1 (delete last paragraph)
        if (paragraphIndex == -1)
        {
            if (paragraphs.Count == 0)
            {
                throw new ArgumentException("無法刪除段落：文檔中沒有段落");
            }
            paragraphIndex = paragraphs.Count - 1;
        }
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }

        var paragraphToDelete = paragraphs[paragraphIndex] as Paragraph;
        if (paragraphToDelete == null)
        {
            throw new InvalidOperationException($"無法獲取索引 {paragraphIndex} 的段落");
        }

        var textPreview = paragraphToDelete.GetText().Trim();
        if (textPreview.Length > 50)
        {
            textPreview = textPreview.Substring(0, 50) + "...";
        }
        
        paragraphToDelete.Remove();

        doc.Save(outputPath);

        var result = $"成功刪除段落 #{paragraphIndex}\n";
        if (!string.IsNullOrEmpty(textPreview))
        {
            result += $"內容預覽: {textPreview}\n";
        }
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    /// Edits paragraph properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional text, formatting options, sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", "paragraphIndex");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

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
                throw new ArgumentException("Cannot edit paragraph: document has no paragraphs. Use insert operation to add paragraphs first.");
            }
        }
        
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count}, valid range: 0-{doc.Sections.Count - 1})");
        
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException($"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count}, valid range: 0-{paragraphs.Count - 1})");
        
        var para = paragraphs[paragraphIndex];
        var builder = new DocumentBuilder(doc);
        
        if (para.FirstChild != null)
        {
            builder.MoveTo(para.FirstChild);
        }
        else
        {
            // Paragraph is empty, move to the paragraph and we'll add content if text parameter is provided
            builder.MoveTo(para);
        }
        
        // Apply font properties
        if (arguments?["fontName"] != null)
        {
            var fontName = arguments["fontName"]?.GetValue<string>();
            builder.Font.Name = fontName ?? "";
        }
        
        if (arguments?["fontNameAscii"] != null)
        {
            var fontNameAscii = arguments["fontNameAscii"]?.GetValue<string>();
            builder.Font.NameAscii = fontNameAscii ?? "";
        }
        
        if (arguments?["fontNameFarEast"] != null)
        {
            var fontNameFarEast = arguments["fontNameFarEast"]?.GetValue<string>();
            builder.Font.NameFarEast = fontNameFarEast ?? "";
        }
        
        if (arguments?["fontSize"] != null)
        {
            var fontSize = arguments["fontSize"]?.GetValue<double>();
            if (fontSize.HasValue)
                builder.Font.Size = fontSize.Value;
        }
        
        if (arguments?["bold"] != null)
        {
            builder.Font.Bold = arguments["bold"]?.GetValue<bool>() ?? false;
        }
        
        if (arguments?["italic"] != null)
        {
            builder.Font.Italic = arguments["italic"]?.GetValue<bool>() ?? false;
        }
        
        if (arguments?["underline"] != null)
        {
            var underline = arguments["underline"]?.GetValue<bool>() ?? false;
            builder.Font.Underline = underline ? Underline.Single : Underline.None;
        }
        
        if (arguments?["color"] != null)
        {
            var colorStr = arguments["color"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(colorStr))
            {
                builder.Font.Color = ColorHelper.ParseColor(colorStr);
            }
        }
        
        // Apply paragraph properties
        var paraFormat = para.ParagraphFormat;
        
        if (arguments?["alignment"] != null)
        {
            var alignment = arguments["alignment"]?.GetValue<string>() ?? "left";
            paraFormat.Alignment = GetAlignment(alignment);
        }
        
        if (arguments?["indentLeft"] != null)
        {
            var indentLeft = arguments["indentLeft"]?.GetValue<double>();
            if (indentLeft.HasValue)
                paraFormat.LeftIndent = indentLeft.Value;
        }
        
        if (arguments?["indentRight"] != null)
        {
            var indentRight = arguments["indentRight"]?.GetValue<double>();
            if (indentRight.HasValue)
                paraFormat.RightIndent = indentRight.Value;
        }
        
        if (arguments?["firstLineIndent"] != null)
        {
            var firstLineIndent = arguments["firstLineIndent"]?.GetValue<double>();
            if (firstLineIndent.HasValue)
                paraFormat.FirstLineIndent = firstLineIndent.Value;
        }
        
        if (arguments?["spaceBefore"] != null)
        {
            var spaceBefore = arguments["spaceBefore"]?.GetValue<double>();
            if (spaceBefore.HasValue)
                paraFormat.SpaceBefore = spaceBefore.Value;
        }
        
        if (arguments?["spaceAfter"] != null)
        {
            var spaceAfter = arguments["spaceAfter"]?.GetValue<double>();
            if (spaceAfter.HasValue)
                paraFormat.SpaceAfter = spaceAfter.Value;
        }
        
        if (arguments?["lineSpacing"] != null || arguments?["lineSpacingRule"] != null)
        {
            var lineSpacing = arguments?["lineSpacing"]?.GetValue<double>();
            var lineSpacingRule = arguments?["lineSpacingRule"]?.GetValue<string>() ?? "single";
            
            var rule = GetLineSpacingRule(lineSpacingRule);
            paraFormat.LineSpacingRule = rule;
            
            if (lineSpacing.HasValue)
            {
                paraFormat.LineSpacing = lineSpacing.Value;
            }
            else if (lineSpacingRule == "single")
            {
                paraFormat.LineSpacing = 12;
            }
            else if (lineSpacingRule == "oneAndHalf")
            {
                paraFormat.LineSpacing = 18;
            }
            else if (lineSpacingRule == "double")
            {
                paraFormat.LineSpacing = 24;
            }
        }
        
        if (arguments?["styleName"] != null)
        {
            var styleName = arguments["styleName"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(styleName))
            {
                try
                {
                    paraFormat.Style = doc.Styles[styleName];
                }
                catch
                {
                    // Style not found, ignore
                }
            }
        }
        
        // Apply tab stops
        if (arguments?["tabStops"] != null)
        {
            var tabStops = arguments["tabStops"]?.AsArray();
            if (tabStops != null && tabStops.Count > 0)
            {
                paraFormat.TabStops.Clear();
                foreach (var ts in tabStops)
                {
                    var position = ts?["position"]?.GetValue<double>() ?? 0;
                    var alignment = ts?["alignment"]?.GetValue<string>() ?? "left";
                    var leader = ts?["leader"]?.GetValue<string>() ?? "none";
                    
                    paraFormat.TabStops.Add(new TabStop(
                        position,
                        GetTabAlignment(alignment),
                        GetTabLeader(leader)
                    ));
                }
            }
        }
        
        var textParam = arguments?["text"]?.GetValue<string>();
        if (!string.IsNullOrEmpty(textParam))
        {
            // Clear existing content and add new text
            para.RemoveAllChildren();
            var newRun = new Run(doc, textParam);
            
            // Apply font settings to the new run
            if (arguments?["fontName"] != null)
                newRun.Font.Name = arguments["fontName"]?.GetValue<string>() ?? "";
            if (arguments?["fontNameAscii"] != null)
                newRun.Font.NameAscii = arguments["fontNameAscii"]?.GetValue<string>() ?? "";
            if (arguments?["fontNameFarEast"] != null)
                newRun.Font.NameFarEast = arguments["fontNameFarEast"]?.GetValue<string>() ?? "";
            if (arguments?["fontSize"] != null)
            {
                var fontSizeValue = arguments["fontSize"]?.GetValue<double?>();
                if (fontSizeValue.HasValue)
                    newRun.Font.Size = fontSizeValue.Value;
            }
            if (arguments?["bold"] != null)
                newRun.Font.Bold = arguments["bold"]?.GetValue<bool>() ?? false;
            if (arguments?["italic"] != null)
                newRun.Font.Italic = arguments["italic"]?.GetValue<bool>() ?? false;
            if (arguments?["underline"] != null)
                newRun.Font.Underline = arguments["underline"]?.GetValue<bool>() == true ? Underline.Single : Underline.None;
            if (arguments?["color"] != null)
            {
                var colorStr = arguments["color"]?.GetValue<string>();
                if (!string.IsNullOrEmpty(colorStr))
                    newRun.Font.Color = ColorHelper.ParseColor(colorStr);
            }
            
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
                foreach (Run run in runs)
                {
                    if (arguments?["fontName"] != null)
                    {
                        var fontName = arguments["fontName"]?.GetValue<string>();
                        run.Font.Name = fontName ?? "";
                    }
                    
                    if (arguments?["fontNameAscii"] != null)
                    {
                        var fontNameAscii = arguments["fontNameAscii"]?.GetValue<string>();
                        run.Font.NameAscii = fontNameAscii ?? "";
                    }
                    
                    if (arguments?["fontNameFarEast"] != null)
                    {
                        var fontNameFarEast = arguments["fontNameFarEast"]?.GetValue<string>();
                        run.Font.NameFarEast = fontNameFarEast ?? "";
                    }
                    
                    if (arguments?["fontSize"] != null)
                    {
                        var fontSize = arguments["fontSize"]?.GetValue<double>();
                        if (fontSize.HasValue)
                            run.Font.Size = fontSize.Value;
                    }
                    
                    if (arguments?["bold"] != null)
                    {
                        run.Font.Bold = arguments["bold"]?.GetValue<bool>() ?? false;
                    }
                    
                    if (arguments?["italic"] != null)
                    {
                        run.Font.Italic = arguments["italic"]?.GetValue<bool>() ?? false;
                    }
                    
                    if (arguments?["underline"] != null)
                    {
                        var underline = arguments["underline"]?.GetValue<bool>() ?? false;
                        run.Font.Underline = underline ? Underline.Single : Underline.None;
                    }
                    
                    if (arguments?["color"] != null)
                    {
                        var colorStr = arguments["color"]?.GetValue<string>();
                        if (!string.IsNullOrEmpty(colorStr))
                        {
                            run.Font.Color = ColorHelper.ParseColor(colorStr);
                        }
                    }
                }
            }
        }
        
        doc.Save(outputPath);
        
        var resultMsg = $"成功編輯段落 {paragraphIndex} 的格式";
        if (!string.IsNullOrEmpty(textParam))
            resultMsg += $"，文字內容已更新";
        return await Task.FromResult(resultMsg);
    }

    /// <summary>
    /// Gets all paragraphs from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional sectionIndex, includeCommentParagraphs, includeTextboxParagraphs</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all paragraphs</returns>
    private async Task<string> GetParagraphsAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var includeEmpty = arguments?["includeEmpty"]?.GetValue<bool?>() ?? true;
        var styleFilter = arguments?["styleFilter"]?.GetValue<string>();
        var includeCommentParagraphs = arguments?["includeCommentParagraphs"]?.GetValue<bool?>() ?? true;
        var includeTextboxParagraphs = arguments?["includeTextboxParagraphs"]?.GetValue<bool?>() ?? true;

        var doc = new Document(path);
        var sb = new StringBuilder();

        List<Paragraph> paragraphs;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }
            // When sectionIndex is specified, only get paragraphs from that section's Body
            paragraphs = doc.Sections[sectionIndex.Value].Body.GetChildNodes(NodeType.Paragraph, includeCommentParagraphs).Cast<Paragraph>().ToList();
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
                foreach (Section section in doc.Sections)
                {
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();
                    paragraphs.AddRange(bodyParagraphs);
                }
            }
        }

        if (!includeEmpty)
        {
            paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.GetText())).ToList();
        }

        if (!string.IsNullOrEmpty(styleFilter))
        {
            paragraphs = paragraphs.Where(p => p.ParagraphFormat.Style?.Name == styleFilter).ToList();
        }
        
        // Filter out textbox paragraphs if includeTextboxParagraphs is false
        if (!includeTextboxParagraphs)
        {
            paragraphs = paragraphs.Where(p =>
            {
                // Check if paragraph is inside a Shape (including nested structures)
                var shapeAncestor = p.GetAncestor(NodeType.Shape);
                if (shapeAncestor != null)
                {
                    var shape = shapeAncestor as Aspose.Words.Drawing.Shape;
                    if (shape != null && shape.ShapeType == Aspose.Words.Drawing.ShapeType.TextBox)
                    {
                        return false; // Exclude textbox paragraphs
                    }
                }
                
                // Also check if paragraph's parent node chain contains a Shape
                // This handles cases where textbox content might be structured differently
                var currentNode = p.ParentNode;
                while (currentNode != null)
                {
                    if (currentNode.NodeType == NodeType.Shape)
                    {
                        var shape = currentNode as Aspose.Words.Drawing.Shape;
                        if (shape != null && shape.ShapeType == Aspose.Words.Drawing.ShapeType.TextBox)
                        {
                            return false; // Exclude textbox paragraphs
                        }
                    }
                    currentNode = currentNode.ParentNode;
                }
                
                return true;
            }).ToList();
        }

        sb.AppendLine($"=== Paragraphs ({paragraphs.Count}) ===");
        if (!includeCommentParagraphs && !includeTextboxParagraphs)
        {
            sb.AppendLine("(Only Body paragraphs - visible in document body, excluding Comment and TextBox paragraphs)");
        }
        else if (!includeCommentParagraphs)
        {
            sb.AppendLine("(Only Body paragraphs - visible in document body, excluding Comment paragraphs)");
        }
        else if (!includeTextboxParagraphs)
        {
            sb.AppendLine("(Includes all paragraphs in document structure, excluding TextBox paragraphs)");
        }
        else
        {
            sb.AppendLine("(Includes all paragraphs in document structure, including Comment objects and TextBoxes)");
        }
        sb.AppendLine();

        for (int i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();
            
            // Determine paragraph location
            string location = "Body";
            string parentNodeInfo = "";
            
            if (para.ParentNode != null)
            {
                var parentNodeType = para.ParentNode.NodeType;
                parentNodeInfo = $"ParentNode: {parentNodeType}";
                
                // Check if paragraph is inside a Comment object
                var commentAncestor = para.GetAncestor(NodeType.Comment);
                if (commentAncestor != null)
                {
                    location = "[Comment]";
                    var comment = commentAncestor as Comment;
                    if (comment != null)
                    {
                        parentNodeInfo += $" (Comment ID: {comment.Id}, Author: {comment.Author})";
                    }
                }
                else
                {
                    // Check if paragraph is inside a TextBox/Shape
                    var shapeAncestor = para.GetAncestor(NodeType.Shape);
                    if (shapeAncestor != null)
                    {
                        var shape = shapeAncestor as Aspose.Words.Drawing.Shape;
                        if (shape != null && shape.ShapeType == Aspose.Words.Drawing.ShapeType.TextBox)
                        {
                            location = "[TextBox]";
                            parentNodeInfo += $" (TextBox)";
                        }
                        else
                        {
                            location = $"[Shape]";
                        }
                    }
                    else
                    {
                        // Check if paragraph is directly in Body
                        var bodyAncestor = para.GetAncestor(NodeType.Body);
                        if (bodyAncestor != null && para.ParentNode.NodeType == NodeType.Body)
                        {
                            location = "Body";
                        }
                        else if (para.ParentNode.NodeType == NodeType.Body)
                        {
                            location = "Body";
                        }
                        else
                        {
                            location = $"[{parentNodeType}]";
                        }
                    }
                }
            }
            
            sb.AppendLine($"[{i}] Location: {location}");
            sb.AppendLine($"    Style: {para.ParagraphFormat.Style?.Name ?? "(none)"}");
            if (!string.IsNullOrEmpty(parentNodeInfo))
            {
                sb.AppendLine($"    {parentNodeInfo}");
            }
            sb.AppendLine($"    Text: {text.Substring(0, Math.Min(100, text.Length))}{(text.Length > 100 ? "..." : "")}");
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    /// Gets format information for a paragraph
    /// </summary>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with paragraph format details</returns>
    private async Task<string> GetParagraphFormatAsync(JsonObject? arguments, string path)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", "paragraphIndex");
        var includeRunDetails = arguments?["includeRunDetails"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null)
        {
            throw new InvalidOperationException($"無法找到索引 {paragraphIndex} 的段落");
        }

        var result = new StringBuilder();
        result.AppendLine($"=== 段落 #{paragraphIndex} 格式資訊 ===\n");

        result.AppendLine("【基本資訊】");
        result.AppendLine($"段落文字: {para.GetText().Trim()}");
        result.AppendLine($"文字長度: {para.GetText().Trim().Length} 字元");
        result.AppendLine($"Run 數量: {para.Runs.Count}");
        result.AppendLine();

        var format = para.ParagraphFormat;
        result.AppendLine("【段落格式】");
        result.AppendLine($"樣式名稱: {format.StyleName}");
        result.AppendLine($"對齊方式: {format.Alignment}");
        result.AppendLine($"左縮排: {format.LeftIndent:F2} pt ({format.LeftIndent / 28.35:F2} cm)");
        result.AppendLine($"右縮排: {format.RightIndent:F2} pt ({format.RightIndent / 28.35:F2} cm)");
        result.AppendLine($"首行縮排: {format.FirstLineIndent:F2} pt ({format.FirstLineIndent / 28.35:F2} cm)");
        result.AppendLine($"段前間距: {format.SpaceBefore:F2} pt");
        result.AppendLine($"段後間距: {format.SpaceAfter:F2} pt");
        result.AppendLine($"行距: {format.LineSpacing:F2} pt");
        result.AppendLine($"行距規則: {format.LineSpacingRule}");
        result.AppendLine();

        if (para.ListFormat != null && para.ListFormat.IsListItem)
        {
            result.AppendLine("【列表格式】");
            result.AppendLine($"是列表項: 是");
            result.AppendLine($"列表層級: {para.ListFormat.ListLevelNumber}");
            if (para.ListFormat.List != null)
            {
                result.AppendLine($"列表 ID: {para.ListFormat.List.ListId}");
            }
            result.AppendLine();
        }

        if (format.Borders.Count > 0)
        {
            result.AppendLine("【邊框】");
            if (format.Borders.Top.LineStyle != LineStyle.None)
                result.AppendLine($"上邊框: {format.Borders.Top.LineStyle}, {format.Borders.Top.LineWidth} pt, 顏色: {format.Borders.Top.Color.Name}");
            if (format.Borders.Bottom.LineStyle != LineStyle.None)
                result.AppendLine($"下邊框: {format.Borders.Bottom.LineStyle}, {format.Borders.Bottom.LineWidth} pt, 顏色: {format.Borders.Bottom.Color.Name}");
            if (format.Borders.Left.LineStyle != LineStyle.None)
                result.AppendLine($"左邊框: {format.Borders.Left.LineStyle}, {format.Borders.Left.LineWidth} pt, 顏色: {format.Borders.Left.Color.Name}");
            if (format.Borders.Right.LineStyle != LineStyle.None)
                result.AppendLine($"右邊框: {format.Borders.Right.LineStyle}, {format.Borders.Right.LineWidth} pt, 顏色: {format.Borders.Right.Color.Name}");
            result.AppendLine();
        }

        if (format.Shading.BackgroundPatternColor.ToArgb() != System.Drawing.Color.Empty.ToArgb())
        {
            result.AppendLine("【背景色】");
            var color = format.Shading.BackgroundPatternColor;
            result.AppendLine($"背景色: #{color.R:X2}{color.G:X2}{color.B:X2}");
            result.AppendLine();
        }

        if (format.TabStops.Count > 0)
        {
            result.AppendLine("【Tab 停駐點】");
            for (int i = 0; i < format.TabStops.Count; i++)
            {
                var tab = format.TabStops[i];
                result.AppendLine($"  Tab {i + 1}: 位置={tab.Position:F2} pt, 對齊={tab.Alignment}, 前導字元={tab.Leader}");
            }
            result.AppendLine();
        }

        if (para.Runs.Count > 0)
        {
            var firstRun = para.Runs[0];
            result.AppendLine("【字型格式（第一個 Run）】");
            
            if (firstRun.Font.NameAscii != firstRun.Font.NameFarEast)
            {
                result.AppendLine($"字體（英文）: {firstRun.Font.NameAscii}");
                result.AppendLine($"字體（中文）: {firstRun.Font.NameFarEast}");
            }
            else
            {
                result.AppendLine($"字體: {firstRun.Font.Name}");
            }
            
            result.AppendLine($"字號: {firstRun.Font.Size} pt");
            
            if (firstRun.Font.Bold) result.AppendLine("粗體: 是");
            if (firstRun.Font.Italic) result.AppendLine("斜體: 是");
            if (firstRun.Font.Underline != Underline.None) result.AppendLine($"底線: {firstRun.Font.Underline}");
            if (firstRun.Font.StrikeThrough) result.AppendLine("刪除線: 是");
            if (firstRun.Font.Superscript) result.AppendLine("上標: 是");
            if (firstRun.Font.Subscript) result.AppendLine("下標: 是");
            
            if (firstRun.Font.Color.ToArgb() != System.Drawing.Color.Empty.ToArgb())
            {
                var color = firstRun.Font.Color;
                result.AppendLine($"顏色: #{color.R:X2}{color.G:X2}{color.B:X2}");
            }
            
            if (firstRun.Font.HighlightColor != System.Drawing.Color.Empty)
            {
                result.AppendLine($"螢光筆: {firstRun.Font.HighlightColor.Name}");
            }
            result.AppendLine();
        }

        if (includeRunDetails && para.Runs.Count > 1)
        {
            result.AppendLine("【Run 詳細資訊】");
            result.AppendLine($"共 {para.Runs.Count} 個 Run:");
            
            for (int i = 0; i < Math.Min(para.Runs.Count, 10); i++)
            {
                var run = para.Runs[i];
                result.AppendLine($"\n  Run #{i}:");
                result.AppendLine($"    文字: {run.Text.Replace("\r", "\\r").Replace("\n", "\\n")}");
                
                if (run.Font.NameAscii != run.Font.NameFarEast)
                {
                    result.AppendLine($"    字體（英文）: {run.Font.NameAscii}");
                    result.AppendLine($"    字體（中文）: {run.Font.NameFarEast}");
                }
                else
                {
                    result.AppendLine($"    字體: {run.Font.Name}");
                }
                
                result.AppendLine($"    字號: {run.Font.Size} pt");
                
                var styles = new List<string>();
                if (run.Font.Bold) styles.Add("粗體");
                if (run.Font.Italic) styles.Add("斜體");
                if (run.Font.Underline != Underline.None) styles.Add($"底線({run.Font.Underline})");
                if (styles.Count > 0)
                    result.AppendLine($"    樣式: {string.Join(", ", styles)}");
            }
            
            if (para.Runs.Count > 10)
            {
                result.AppendLine($"\n  ... 還有 {para.Runs.Count - 10} 個 Run（已省略）");
            }
            result.AppendLine();
        }

        result.AppendLine("【JSON 格式（可用於 word_edit_paragraph）】");
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
    /// Copies paragraph format from source to target
    /// </summary>
    /// <param name="arguments">JSON arguments containing sourceIndex, targetIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> CopyParagraphFormatAsync(JsonObject? arguments, string path, string outputPath)
    {
        var sourceParagraphIndex = ArgumentHelper.GetInt(arguments, "sourceParagraphIndex", "sourceParagraphIndex");
        var targetParagraphIndex = ArgumentHelper.GetInt(arguments, "targetParagraphIndex", "targetParagraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (sourceParagraphIndex < 0 || sourceParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"來源段落索引 {sourceParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (targetParagraphIndex < 0 || targetParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"目標段落索引 {targetParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var sourcePara = paragraphs[sourceParagraphIndex] as Paragraph;
        var targetPara = paragraphs[targetParagraphIndex] as Paragraph;
        
        if (sourcePara == null || targetPara == null)
        {
            throw new InvalidOperationException("無法獲取段落");
        }
        
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
        for (int i = 0; i < sourcePara.ParagraphFormat.TabStops.Count; i++)
        {
            var tabStop = sourcePara.ParagraphFormat.TabStops[i];
            targetPara.ParagraphFormat.TabStops.Add(tabStop.Position, tabStop.Alignment, tabStop.Leader);
        }
        
        doc.Save(outputPath);
        
        var result = $"成功複製段落格式\n";
        result += $"來源段落: #{sourceParagraphIndex}\n";
        result += $"目標段落: #{targetParagraphIndex}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Merges multiple paragraphs into one
    /// </summary>
    /// <param name="arguments">JSON arguments containing startIndex, endIndex, optional sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> MergeParagraphsAsync(JsonObject? arguments, string path, string outputPath)
    {
        var startParagraphIndex = ArgumentHelper.GetInt(arguments, "startParagraphIndex", "startParagraphIndex");
        var endParagraphIndex = ArgumentHelper.GetInt(arguments, "endParagraphIndex", "endParagraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"起始段落索引 {startParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"結束段落索引 {endParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (startParagraphIndex > endParagraphIndex)
        {
            throw new ArgumentException($"起始段落索引 {startParagraphIndex} 不能大於結束段落索引 {endParagraphIndex}");
        }
        
        if (startParagraphIndex == endParagraphIndex)
        {
            throw new ArgumentException("起始和結束段落索引相同，無需合併");
        }
        
        var startPara = paragraphs[startParagraphIndex] as Paragraph;
        if (startPara == null)
        {
            throw new InvalidOperationException("無法獲取起始段落");
        }
        
        for (int i = startParagraphIndex + 1; i <= endParagraphIndex; i++)
        {
            var para = paragraphs[i] as Paragraph;
            if (para != null)
            {
                if (startPara.Runs.Count > 0)
                {
                    var spaceRun = new Run(doc, " ");
                    startPara.AppendChild(spaceRun);
                }
                
                var runsToMove = para.Runs.ToArray();
                foreach (var run in runsToMove)
                {
                    startPara.AppendChild(run);
                }
                
                para.Remove();
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功合併段落\n";
        result += $"合併範圍: 段落 #{startParagraphIndex} 到 #{endParagraphIndex}\n";
        result += $"合併段落數: {endParagraphIndex - startParagraphIndex + 1}\n";
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";
        
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

