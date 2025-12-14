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
    public string Description => @"Manage lists in Word documents. Supports 6 operations: add_list, add_item, delete_item, edit_item, set_format, get_format.

Usage examples:
- Add bullet list: word_list(path='doc.docx', items=['Item 1', 'Item 2', 'Item 3'])
- Add numbered list: word_list(path='doc.docx', items=['First', 'Second'], listType='number')
- Add list item: word_list(path='doc.docx', text='New item', styleName='!標題4-數字')
- Delete list item: word_list(path='doc.docx', paragraphIndex=0)
- Edit list item: word_list(path='doc.docx', paragraphIndex=0, text='Updated text')
- Get list format: word_list(path='doc.docx', paragraphIndex=0)

Note: The 'operation' parameter is optional and will be auto-inferred from other parameters. You can also explicitly specify it.";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add_list': Add a new list (required params: path, items)
- 'add_item': Add an item to existing list (required params: path, text, styleName)
- 'delete_item': Delete a list item (required params: path, paragraphIndex)
- 'edit_item': Edit a list item (required params: path, paragraphIndex, text)
- 'set_format': Set list format (required params: path, paragraphIndex)
- 'get_format': Get list format (required params: path, paragraphIndex). Note: This operation can only be used on list item paragraphs. If the paragraph is not a list item, it will return a message indicating that the paragraph is not a list item.",
                @enum = new[] { "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format" }
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
            // Add list parameters
            items = new
            {
                type = "array",
                description = @"List items for add_list operation.
Format: Array of strings.
Simple format: ['Item 1', 'Item 2', 'Item 3']",
                items = new { type = "string" }
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
                description = "List item text content (required for add_item and edit_item operations)"
            },
            styleName = new
            {
                type = "string",
                description = "Style name for the list item (required for add_item operation). Example: '!標題4-數字'. Use word_get_styles tool to see available styles."
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
                description = "Paragraph index (0-based, required for delete_item, edit_item, set_format, and get_format operations). Note: For get_format operation, this must be a list item paragraph. If the paragraph is not a list item, the operation will return a message indicating that the paragraph is not a list item."
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
        if (arguments == null)
        {
            throw new ArgumentException("❌ Arguments cannot be null\n\n" +
                                      $"📝 Usage example: word_list(path='doc.docx', items=['Item 1', 'Item 2'])");
        }
        
        if (!arguments.ContainsKey("path"))
        {
            var providedKeys = arguments.Select(kvp => kvp.Key).ToList();
            throw new ArgumentException($"❌ Required parameter 'path' is missing\n\n" +
                                      $"📋 Provided parameters: {(providedKeys.Count > 0 ? string.Join(", ", providedKeys.Select(k => $"'{k}'")) : "none")}\n\n" +
                                      $"📝 Usage examples:\n" +
                                      $"  word_list(path='doc.docx', items=['Item 1', 'Item 2', 'Item 3'])\n" +
                                      $"  word_list(path='doc.docx', text='New item', styleName='!標題4-數字')\n" +
                                      $"  word_list(path='doc.docx', paragraphIndex=0)\n\n" +
                                      $"💡 Note: 'path' parameter is required for all operations.");
        }
        
        var pathValue = arguments["path"];
        if (pathValue == null)
        {
            throw new ArgumentException("❌ Parameter 'path' is null\n\n" +
                                      $"📝 Usage example: word_list(path='doc.docx', items=['Item 1', 'Item 2'])\n\n" +
                                      $"💡 Note: 'path' must be a non-null string value.");
        }
        
        string path;
        try
        {
            path = pathValue.GetValue<string>();
        }
        catch (Exception ex)
        {
            var pathType = pathValue.GetType().Name;
            throw new ArgumentException($"❌ Parameter 'path' has incorrect type\n\n" +
                                      $"📋 Current type: {pathType}\n" +
                                      $"📋 Current value: {pathValue}\n\n" +
                                      $"📝 Expected: string (e.g., 'doc.docx')\n\n" +
                                      $"💡 Error: {ex.Message}");
        }
        
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("❌ Parameter 'path' cannot be empty\n\n" +
                                      $"📝 Usage example: word_list(path='doc.docx', items=['Item 1', 'Item 2'])\n\n" +
                                      $"💡 Note: 'path' must be a non-empty string containing the document file path.");
        }
        
        SecurityHelper.ValidateFilePath(path, "path");
        
        // Auto-infer operation if not provided
        string operation;
        if (!arguments.ContainsKey("operation") || arguments["operation"] == null)
        {
            // Auto-infer operation from provided parameters
            // This allows users to call word_list without explicitly specifying operation
            var providedKeys = arguments.Select(kvp => kvp.Key).ToList();
            var providedParamsInfo = $"提供的參數: {string.Join(", ", providedKeys.Select(k => $"'{k}'"))}";
            
            // Infer operation based on provided parameters
            if (arguments.ContainsKey("items") && arguments["items"] != null)
            {
                // Has items parameter -> add_list
                operation = "add_list";
            }
            else if (arguments.ContainsKey("text") && arguments["text"] != null)
            {
                if (arguments.ContainsKey("itemIndex") && arguments["itemIndex"] != null)
                {
                    // Has text and itemIndex -> edit_item
                    operation = "edit_item";
                }
                else
                {
                    // Has text but no itemIndex -> add_item
                    operation = "add_item";
                }
            }
            else if (arguments.ContainsKey("itemIndex") && arguments["itemIndex"] != null)
            {
                if (arguments.ContainsKey("alignment") || arguments.ContainsKey("leftIndent") || 
                    arguments.ContainsKey("firstLineIndent") || arguments.ContainsKey("spaceAfter"))
                {
                    // Has itemIndex and format parameters -> set_format
                    operation = "set_format";
                }
                else
                {
                    // Has itemIndex but no text -> delete_item (or get_format)
                    // Check if it's a read operation (no outputPath or outputPath == path)
                    var docPath = arguments["path"]?.GetValue<string>();
                    var docOutputPath = arguments["outputPath"]?.GetValue<string>() ?? docPath;
                    if (docPath == docOutputPath && !arguments.ContainsKey("text"))
                    {
                        // Same path and no text -> get_format (read operation)
                        operation = "get_format";
                    }
                    else
                    {
                        // Different path or has text -> delete_item
                        operation = "delete_item";
                    }
                }
            }
            else
            {
                // Cannot infer operation
                var availableOps = new[] { "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format" };
                throw new ArgumentException($"❌ 缺少必需參數 'operation'，且無法從提供的參數自動推斷操作類型\n\n" +
                                          $"📋 {providedParamsInfo}\n\n" +
                                          $"📋 可用操作: {string.Join(", ", availableOps)}\n\n" +
                                          $"📝 使用範例:\n" +
                                          $"  1. 添加項目符號列表（自動推斷）:\n" +
                                          $"     word_list(path='doc.docx', items=['項目1', '項目2', '項目3'])\n\n" +
                                          $"  2. 添加編號列表（自動推斷）:\n" +
                                          $"     word_list(path='doc.docx', items=['第一項', '第二項'], listType='number')\n\n" +
                                          $"  3. 添加列表項目（自動推斷）:\n" +
                                          $"     word_list(path='doc.docx', text='新項目')\n\n" +
                                          $"  4. 刪除列表項目（明確指定）:\n" +
                                          $"     word_list(operation='delete_item', path='doc.docx', itemIndex=0)\n\n" +
                                          $"  5. 編輯列表項目（自動推斷）:\n" +
                                          $"     word_list(path='doc.docx', itemIndex=0, text='修改後的文字')\n\n" +
                                          $"  6. 獲取列表格式（自動推斷）:\n" +
                                          $"     word_list(path='doc.docx', itemIndex=0)\n\n" +
                                          $"💡 提示: 如果自動推斷失敗，請明確指定 operation 參數");
            }
            
            // Add the inferred operation to arguments for consistency
            arguments["operation"] = operation;
        }
        else
        {
            operation = arguments["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
            
            // Validate operation value
            var validOperations = new[] { "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format" };
            if (!validOperations.Contains(operation))
            {
                throw new ArgumentException($"Invalid operation: '{operation}'. Valid operations: {string.Join(", ", validOperations.Select(op => $"'{op}'"))}");
            }
        }
        
        var outputPath = arguments["outputPath"]?.GetValue<string>() ?? path;
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

    /// <summary>
    /// Adds a list to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing items array, optional listType, listStyle, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddListAsync(JsonObject? arguments, string path, string outputPath)
    {
        var items = arguments?["items"];
        if (items == null)
        {
            throw new ArgumentException("❌ items parameter is required");
        }
        
        try
        {
            var parsedItems = ParseItems(items);
            var listType = arguments?["listType"]?.GetValue<string>() ?? "bullet";
            var bulletChar = arguments?["bulletChar"]?.GetValue<string>() ?? "●";
            var numberFormat = arguments?["numberFormat"]?.GetValue<string>() ?? "arabic";
            
            // Open document and create list
            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            
            // Create list
            var list = doc.Lists.Add(listType == "number" ? ListTemplate.NumberDefault : ListTemplate.BulletDefault);
            
            // Configure list format
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
            
            // Add list items
            foreach (var item in parsedItems)
            {
                builder.ListFormat.List = list;
                builder.ListFormat.ListLevelNumber = Math.Min(item.level, 8);
                builder.Writeln(item.text);
            }
            
            // Remove list formatting after adding items
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
        catch (Exception ex)
        {
            throw new ArgumentException($"❌ Error creating list: {ex.Message}");
        }
    }

    /// <summary>
    /// Adds an item to an existing list
    /// </summary>
    /// <param name="arguments">JSON arguments containing listIndex, text, optional insertAt, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddListItemAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = ArgumentHelper.GetString(arguments, "text", "text");
        var styleName = ArgumentHelper.GetString(arguments, "styleName", "styleName");
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

    /// <summary>
    /// Deletes an item from a list
    /// </summary>
    /// <param name="arguments">JSON arguments containing listIndex, itemIndex, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteListItemAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", "paragraphIndex");

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

    /// <summary>
    /// Edits a list item
    /// </summary>
    /// <param name="arguments">JSON arguments containing listIndex, itemIndex, text, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditListItemAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", "paragraphIndex");
        var text = ArgumentHelper.GetString(arguments, "text", "text");
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

    /// <summary>
    /// Sets list format properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing listIndex, optional listType, listStyle, formatting options</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> SetListFormatAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", "paragraphIndex");
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


    /// <summary>
    /// Gets list format information
    /// </summary>
    /// <param name="arguments">JSON arguments containing listIndex</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with list format details</returns>
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
            result.AppendLine($"說明: 此段落不是列表項，無法獲取列表格式資訊。如需將此段落轉換為列表項，請使用 insert_list 或 set_list_style 操作");
        }
    }

    private List<(string text, int level)> ParseItems(JsonNode? itemsNode)
    {
        var items = new List<(string text, int level)>();

        if (itemsNode == null)
        {
            throw new ArgumentException("❌ items 參數不能為 null\n\n" +
                                      $"📝 請提供一個數組，格式:\n" +
                                      $"  簡單格式: [\"項目1\", \"項目2\", \"項目3\"]\n" +
                                      $"  帶級別格式: [{{\"text\": \"項目1\", \"level\": 0}}, {{\"text\": \"子項目\", \"level\": 1}}]");
        }

        try
        {
            var itemsArray = itemsNode.AsArray();
            if (itemsArray == null)
            {
                var nodeType = itemsNode.GetType().Name;
                var nodeValue = itemsNode.ToString();
                throw new ArgumentException($"❌ items 參數必須是一個數組\n\n" +
                                          $"📋 當前類型: {nodeType}\n" +
                                          $"📋 當前值: {nodeValue}\n\n" +
                                          $"📝 正確格式範例:\n" +
                                          $"  簡單格式: [\"項目1\", \"項目2\", \"項目3\"]\n" +
                                          $"  帶級別格式: [{{\"text\": \"項目1\", \"level\": 0}}, {{\"text\": \"子項目\", \"level\": 1}}]");
            }
            
            if (itemsArray.Count == 0)
            {
                throw new ArgumentException("❌ items 數組不能為空\n\n" +
                                          $"📝 請至少提供一個項目，例如: [\"項目1\"]");
            }
            
            foreach (var item in itemsArray)
            {
                if (item == null)
                {
                    continue; // Skip null items
                }
                
                if (item is JsonValue jsonValue)
                {
                    // Simple string item
                    try
                    {
                        var text = jsonValue.GetValue<string>();
                        if (!string.IsNullOrEmpty(text))
                        {
                            items.Add((text, 0));
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException($"❌ 無法解析列表項目為字符串: {ex.Message}\n\n" +
                                                  $"📋 項目值: {item}\n\n" +
                                                  $"📝 正確格式: 字符串，例如 \"項目1\"");
                    }
                }
                else if (item is JsonObject jsonObj)
                {
                    // Object with text and level
                    var text = jsonObj["text"]?.GetValue<string>();
                    if (string.IsNullOrEmpty(text))
                    {
                        var objKeys = string.Join(", ", jsonObj.Select(kvp => $"'{kvp.Key}'"));
                        throw new ArgumentException($"❌ 列表項目對象必須包含 'text' 屬性\n\n" +
                                                  $"📋 當前對象的鍵: {objKeys}\n\n" +
                                                  $"📝 正確格式: {{\"text\": \"項目文字\", \"level\": 0}}");
                    }
                    
                    var level = jsonObj["level"]?.GetValue<int>() ?? 0;
                    if (level < 0 || level > 8)
                    {
                        level = Math.Max(0, Math.Min(8, level)); // Clamp to valid range
                    }
                    
                    items.Add((text, level));
                }
                else
                {
                    throw new ArgumentException($"❌ 無效的列表項目格式\n\n" +
                                              $"📋 項目類型: {item.GetType().Name}\n" +
                                              $"📋 項目值: {item}\n\n" +
                                              $"📝 正確格式:\n" +
                                              $"  字符串: \"項目文字\"\n" +
                                              $"  對象: {{\"text\": \"項目文字\", \"level\": 0}}");
                }
            }
            
            if (items.Count == 0)
            {
                throw new ArgumentException("❌ 解析後沒有有效的列表項目\n\n" +
                                          $"📝 請確保 items 數組包含至少一個有效的字符串或對象");
            }
        }
        catch (ArgumentException)
        {
            throw; // Re-throw ArgumentException as-is
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"❌ 解析 items 參數時發生錯誤: {ex.Message}\n\n" +
                                      $"📋 錯誤類型: {ex.GetType().Name}\n\n" +
                                      $"📝 請確保 items 是一個數組，格式:\n" +
                                      $"  簡單格式: [\"項目1\", \"項目2\"]\n" +
                                      $"  帶級別格式: [{{\"text\": \"項目1\", \"level\": 0}}, ...]", ex);
        }

        if (items.Count == 0)
        {
            throw new ArgumentException("無法解析任何有效的列表項目。請檢查 items 參數格式");
        }

        return items;
    }
}

