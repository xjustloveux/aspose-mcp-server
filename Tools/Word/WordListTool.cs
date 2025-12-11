using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for list operations in Word documents
/// Merges: WordAddListTool, WordAddListItemTool, WordDeleteListItemTool, WordEditListItemTool,
/// WordSetListFormatTool, WordGetListFormatTool
/// </summary>
public class WordListTool : IAsposeTool
{
    public string Description => "Manage lists in Word documents: add list, add item, delete item, edit item, set format, get format";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add_list', 'add_item', 'delete_item', 'edit_item', 'set_format', 'get_format'",
                @enum = new[] { "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for write operations)"
            },
            // Add list parameters
            items = new
            {
                type = "array",
                description = "List items (strings or objects with 'text' and 'level', required for add_list operation)",
                items = new
                {
                    oneOf = new object[]
                    {
                        new { type = "string" },
                        new
                        {
                            type = "object",
                            properties = new
                            {
                                text = new { type = "string" },
                                level = new { type = "number", description = "Indent level (0-8, default: 0)" }
                            }
                        }
                    }
                }
            },
            listType = new
            {
                type = "string",
                description = "List type: bullet, number, custom (optional, default: bullet, for add_list operation)",
                @enum = new[] { "bullet", "number", "custom" }
            },
            bulletChar = new
            {
                type = "string",
                description = "Custom bullet character (optional, for custom type, e.g., '●', '■', '▪')"
            },
            numberFormat = new
            {
                type = "string",
                description = "Number format for numbered lists: arabic, roman, letter (optional, default: arabic, for add_list operation)",
                @enum = new[] { "arabic", "roman", "letter" }
            },
            // Add item parameters
            text = new
            {
                type = "string",
                description = "List item text content (required for add_item/edit_item operations)"
            },
            styleName = new
            {
                type = "string",
                description = "Style name for the list item (required for add_item operation, e.g., '!標題4-數字')"
            },
            listLevel = new
            {
                type = "number",
                description = "List level (0-8, optional, for add_item operation)"
            },
            applyStyleIndent = new
            {
                type = "boolean",
                description = "If true, uses the indentation defined in the style (optional, default: true, for add_item operation)"
            },
            // Delete/Edit item parameters
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, required for delete_item/edit_item/set_format operations, optional for get_format operation)"
            },
            level = new
            {
                type = "number",
                description = "List level (0-8, optional, for edit_item operation)"
            },
            // Set format parameters
            numberStyle = new
            {
                type = "string",
                description = "Number style: arabic, roman, letter, bullet, none (optional, for set_format operation)",
                @enum = new[] { "arabic", "roman", "letter", "bullet", "none" }
            },
            indentLevel = new
            {
                type = "number",
                description = "Indentation level (0-8, optional, for set_format operation). Each level = 36 points (0.5 inch)"
            },
            leftIndent = new
            {
                type = "number",
                description = "Left indent in points (optional, overrides indentLevel if provided, for set_format operation)"
            },
            firstLineIndent = new
            {
                type = "number",
                description = "First line indent in points (optional, negative for hanging indent, for set_format operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        return operation switch
        {
            "add_list" => await AddListAsync(arguments, path, outputPath),
            "add_item" => await AddListItemAsync(arguments, path, outputPath),
            "delete_item" => await DeleteListItemAsync(arguments, path, outputPath),
            "edit_item" => await EditListItemAsync(arguments, path, outputPath),
            "set_format" => await SetListFormatAsync(arguments, path, outputPath),
            "get_format" => await GetListFormatAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddListAsync(JsonObject? arguments, string path, string outputPath)
    {
        var items = arguments?["items"] ?? throw new ArgumentException("items is required");
        var listType = arguments?["listType"]?.GetValue<string>() ?? "bullet";
        var bulletChar = arguments?["bulletChar"]?.GetValue<string>() ?? "●";
        var numberFormat = arguments?["numberFormat"]?.GetValue<string>() ?? "arabic";

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var parsedItems = ParseItems(items);
        if (parsedItems.Count == 0)
        {
            throw new ArgumentException("items 不能為空");
        }

        var list = doc.Lists.Add(listType == "number" ? ListTemplate.NumberDefault : ListTemplate.BulletDefault);

        if (listType == "custom" && !string.IsNullOrEmpty(bulletChar))
        {
            list.ListLevels[0].NumberFormat = bulletChar;
            list.ListLevels[0].NumberStyle = NumberStyle.Bullet;
        }
        else if (listType == "number")
        {
            var numStyle = numberFormat.ToLower() switch
            {
                "roman" => NumberStyle.UppercaseRoman,
                "letter" => NumberStyle.UppercaseLetter,
                _ => NumberStyle.Arabic
            };

            for (int i = 0; i < list.ListLevels.Count; i++)
            {
                list.ListLevels[i].NumberStyle = numStyle;
            }
        }

        foreach (var item in parsedItems)
        {
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = Math.Min(item.level, 8);
            builder.Writeln(item.text);
        }

        builder.ListFormat.RemoveNumbers();
        doc.Save(outputPath);

        var result = $"成功添加清單\n";
        result += $"類型: {listType}\n";
        if (listType == "custom") result += $"項目符號: {bulletChar}\n";
        if (listType == "number") result += $"數字格式: {numberFormat}\n";
        result += $"項目數: {parsedItems.Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> AddListItemAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var styleName = arguments?["styleName"]?.GetValue<string>() ?? throw new ArgumentException("styleName is required");
        var listLevel = arguments?["listLevel"]?.GetValue<int>() ?? 0;
        var applyStyleIndent = arguments?["applyStyleIndent"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var style = doc.Styles[styleName];
        if (style == null)
        {
            throw new ArgumentException($"找不到樣式 '{styleName}'，可用樣式請使用 word_get_styles 工具查看");
        }

        var para = new Paragraph(doc);
        para.ParagraphFormat.StyleName = styleName;

        if (!applyStyleIndent && listLevel > 0)
        {
            para.ParagraphFormat.LeftIndent = listLevel * 36;
        }

        var run = new Run(doc, text);
        para.AppendChild(run);
        builder.CurrentParagraph.ParentNode.AppendChild(para);

        doc.Save(outputPath);

        var result = "成功添加清單項目\n";
        result += $"樣式: {styleName}\n";
        result += $"級別: {listLevel}\n";
        
        if (applyStyleIndent)
        {
            result += "縮排: 使用樣式定義的縮排（推薦）\n";
        }
        else if (listLevel > 0)
        {
            result += $"縮排: 手動設定 ({listLevel * 36} points)\n";
        }
        
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> DeleteListItemAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var paraToDelete = paragraphs[paragraphIndex] as Paragraph;
        if (paraToDelete == null)
        {
            throw new InvalidOperationException($"無法獲取索引 {paragraphIndex} 的段落");
        }
        
        string itemText = paraToDelete.GetText().Trim();
        string itemPreview = itemText.Length > 50 ? itemText.Substring(0, 50) + "..." : itemText;
        bool isListItem = paraToDelete.ListFormat.IsListItem;
        string listInfo = isListItem ? "（清單項目）" : "（一般段落）";
        
        paraToDelete.Remove();
        doc.Save(outputPath);
        
        var result = $"成功刪除清單項目 #{paragraphIndex}{listInfo}\n";
        if (!string.IsNullOrEmpty(itemPreview))
        {
            result += $"內容預覽: {itemPreview}\n";
        }
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> EditListItemAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var level = arguments?["level"]?.GetValue<int?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null)
        {
            throw new InvalidOperationException($"無法獲取索引 {paragraphIndex} 的段落");
        }
        
        para.Runs.Clear();
        var run = new Run(doc, text);
        para.AppendChild(run);
        
        if (level.HasValue && level.Value >= 0 && level.Value <= 8)
        {
            para.ParagraphFormat.LeftIndent = level.Value * 36;
        }
        
        doc.Save(outputPath);
        
        var result = $"成功編輯清單項目\n";
        result += $"段落索引: {paragraphIndex}\n";
        result += $"新文字: {text}\n";
        if (level.HasValue)
        {
            result += $"級別: {level.Value}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> SetListFormatAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var numberStyle = arguments?["numberStyle"]?.GetValue<string>();
        var indentLevel = arguments?["indentLevel"]?.GetValue<int?>();
        var leftIndent = arguments?["leftIndent"]?.GetValue<double?>();
        var firstLineIndent = arguments?["firstLineIndent"]?.GetValue<double?>();

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
        
        var changes = new List<string>();
        
        if (!string.IsNullOrEmpty(numberStyle) && para.ListFormat.IsListItem)
        {
            var list = para.ListFormat.List;
            if (list != null)
            {
                var level = para.ListFormat.ListLevelNumber;
                var listLevel = list.ListLevels[level];
                
                var style = numberStyle.ToLower() switch
                {
                    "arabic" => NumberStyle.Arabic,
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    "bullet" => NumberStyle.Bullet,
                    "none" => NumberStyle.None,
                    _ => NumberStyle.Arabic
                };
                
                listLevel.NumberStyle = style;
                changes.Add($"編號樣式: {numberStyle}");
            }
        }
        
        if (indentLevel.HasValue)
        {
            para.ParagraphFormat.LeftIndent = indentLevel.Value * 36;
            changes.Add($"縮排層級: {indentLevel.Value}");
        }
        
        if (leftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = leftIndent.Value;
            changes.Add($"左縮排: {leftIndent.Value} 點");
        }
        
        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
            changes.Add($"首行縮排: {firstLineIndent.Value} 點");
        }
        
        doc.Save(outputPath);
        
        var result = $"成功設定清單格式\n";
        result += $"段落索引: {paragraphIndex}\n";
        if (changes.Count > 0)
        {
            result += $"變更內容: {string.Join("、", changes)}\n";
        }
        else
        {
            result += "未提供變更參數\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> GetListFormatAsync(JsonObject? arguments, string path)
    {
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var result = new StringBuilder();

        result.AppendLine("=== 文檔列表格式資訊 ===\n");

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
            
            var para = paragraphs[paragraphIndex.Value];
            AppendListFormatInfo(result, para, paragraphIndex.Value);
        }
        else
        {
            var listParagraphs = paragraphs
                .Where(p => p.ListFormat != null && p.ListFormat.IsListItem)
                .ToList();
            
            result.AppendLine($"總列表段落數: {listParagraphs.Count}\n");
            
            if (listParagraphs.Count == 0)
            {
                result.AppendLine("未找到列表段落");
                return await Task.FromResult(result.ToString());
            }
            
            for (int i = 0; i < listParagraphs.Count; i++)
            {
                var para = listParagraphs[i];
                var paraIndex = paragraphs.IndexOf(para);
                AppendListFormatInfo(result, para, paraIndex);
                if (i < listParagraphs.Count - 1)
                {
                    result.AppendLine();
                }
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendListFormatInfo(StringBuilder result, Paragraph para, int paraIndex)
    {
        result.AppendLine($"【段落 {paraIndex}】");
        result.AppendLine($"內容預覽: {para.GetText().Trim().Substring(0, Math.Min(50, para.GetText().Trim().Length))}...");
        
        if (para.ListFormat != null && para.ListFormat.IsListItem)
        {
            result.AppendLine($"是否列表項: 是");
            result.AppendLine($"列表級別: {para.ListFormat.ListLevelNumber}");
            
            if (para.ListFormat.List != null)
            {
                result.AppendLine($"列表ID: {para.ListFormat.List.ListId}");
            }
            
            if (para.ListFormat.ListLevel != null)
            {
                var level = para.ListFormat.ListLevel;
                result.AppendLine($"列表符號: {level.NumberFormat}");
                result.AppendLine($"對齊方式: {level.Alignment}");
                result.AppendLine($"文本位置: {level.TextPosition}");
                result.AppendLine($"編號樣式: {level.NumberStyle}");
            }
        }
        else
        {
            result.AppendLine($"是否列表項: 否");
        }
    }

    private List<(string text, int level)> ParseItems(JsonNode? itemsNode)
    {
        var items = new List<(string text, int level)>();

        if (itemsNode == null)
            return items;

        try
        {
            var itemsArray = itemsNode.AsArray();
            foreach (var item in itemsArray)
            {
                if (item is JsonValue jsonValue)
                {
                    items.Add((jsonValue.GetValue<string>(), 0));
                }
                else if (item is JsonObject jsonObj)
                {
                    var text = jsonObj["text"]?.GetValue<string>() ?? "";
                    var level = jsonObj["level"]?.GetValue<int>() ?? 0;
                    items.Add((text, level));
                }
            }
        }
        catch
        {
            // Return empty list on parse error
        }

        return items;
    }
}

