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
- Add list item: word_list(path='doc.docx', text='New item', styleName='Heading 4')
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
                description = "Style name for the list item (required for add_item operation). Example: 'Heading 4'. Use word_get_styles tool to see available styles."
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
                                      $"  word_list(path='doc.docx', text='New item', styleName='Heading 4')\n" +
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
        
        SecurityHelper.ValidateFilePath(path);
        
        // Auto-infer operation if not provided
        string operation;
        if (!arguments.ContainsKey("operation") || arguments["operation"] == null)
        {
            // Auto-infer operation from provided parameters
            // This allows users to call word_list without explicitly specifying operation
            var providedKeys = arguments.Select(kvp => kvp.Key).ToList();
            var providedParamsInfo = $"Provided parameters: {string.Join(", ", providedKeys.Select(k => $"'{k}'"))}";
            
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
                    var docPath = ArgumentHelper.GetStringNullable(arguments, "path");
                    var docOutputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? docPath;
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
                throw new ArgumentException($"❌ Required parameter 'operation' is missing and cannot be inferred from provided parameters\n\n" +
                                          $"📋 {providedParamsInfo}\n\n" +
                                          $"📋 Available operations: {string.Join(", ", availableOps)}\n\n" +
                                          $"📝 Usage examples:\n" +
                                          $"  1. Add bullet list (auto-inferred):\n" +
                                          $"     word_list(path='doc.docx', items=['Item 1', 'Item 2', 'Item 3'])\n\n" +
                                          $"  2. Add numbered list (auto-inferred):\n" +
                                          $"     word_list(path='doc.docx', items=['First item', 'Second item'], listType='number')\n\n" +
                                          $"  3. Add list item (auto-inferred):\n" +
                                          $"     word_list(path='doc.docx', text='New item')\n\n" +
                                          $"  4. Delete list item (explicit):\n" +
                                          $"     word_list(operation='delete_item', path='doc.docx', itemIndex=0)\n\n" +
                                          $"  5. Edit list item (auto-inferred):\n" +
                                          $"     word_list(path='doc.docx', itemIndex=0, text='Modified text')\n\n" +
                                          $"  6. Get list format (auto-inferred):\n" +
                                          $"     word_list(path='doc.docx', itemIndex=0)\n\n" +
                                          $"💡 Tip: If auto-inference fails, explicitly specify the operation parameter");
            }
            
            // Add the inferred operation to arguments for consistency
            arguments["operation"] = operation;
        }
        else
        {
            operation = ArgumentHelper.GetString(arguments, "operation");
            
            // Validate operation value
            var validOperations = new[] { "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format" };
            if (!validOperations.Contains(operation))
            {
                throw new ArgumentException($"Invalid operation: '{operation}'. Valid operations: {string.Join(", ", validOperations.Select(op => $"'{op}'"))}");
            }
        }
        
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
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
            var listType = ArgumentHelper.GetString(arguments, "listType", "bullet");
            var bulletChar = ArgumentHelper.GetString(arguments, "bulletChar", "●");
            var numberFormat = ArgumentHelper.GetString(arguments, "numberFormat", "arabic");
            
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
            
            var result = $"List added successfully\n";
            result += $"Type: {listType}\n";
            if (listType == "custom") result += $"Bullet character: {bulletChar}\n";
            if (listType == "number") result += $"Number format: {numberFormat}\n";
            result += $"Item count: {parsedItems.Count}\n";
            result += $"Output: {outputPath}";

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
        var text = ArgumentHelper.GetString(arguments, "text");
        var styleName = ArgumentHelper.GetString(arguments, "styleName");
        var listLevel = ArgumentHelper.GetInt(arguments, "listLevel", 0);
        var applyStyleIndent = ArgumentHelper.GetBool(arguments, "applyStyleIndent");

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var style = doc.Styles[styleName];
        if (style == null)
        {
            throw new ArgumentException($"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
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

        var result = "List item added successfully\n";
        result += $"Style: {styleName}\n";
        result += $"Level: {listLevel}\n";
        
        if (applyStyleIndent)
        {
            result += "Indent: Using style-defined indent (recommended)\n";
        }
        else if (listLevel > 0)
        {
            result += $"Indent: Manually set ({listLevel * 36} points)\n";
        }
        
        result += $"Output: {outputPath}";

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
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");
        }
        
        var paraToDelete = paragraphs[paragraphIndex] as Paragraph;
        if (paraToDelete == null)
        {
            throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex}");
        }
        
        string itemText = paraToDelete.GetText().Trim();
        string itemPreview = itemText.Length > 50 ? itemText.Substring(0, 50) + "..." : itemText;
        bool isListItem = paraToDelete.ListFormat.IsListItem;
        string listInfo = isListItem ? " (list item)" : " (regular paragraph)";
        
        paraToDelete.Remove();
        doc.Save(outputPath);
        
        var result = $"List item #{paragraphIndex} deleted successfully{listInfo}\n";
        if (!string.IsNullOrEmpty(itemPreview))
        {
            result += $"Content preview: {itemPreview}\n";
        }
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"Output: {outputPath}";
        
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
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
        var text = ArgumentHelper.GetString(arguments, "text");
        var level = ArgumentHelper.GetIntNullable(arguments, "level");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");
        }
        
        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null)
        {
            throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex}");
        }
        
        para.Runs.Clear();
        var run = new Run(doc, text);
        para.AppendChild(run);
        
        if (level.HasValue && level.Value >= 0 && level.Value <= 8)
        {
            para.ParagraphFormat.LeftIndent = level.Value * 36;
        }
        
        doc.Save(outputPath);
        
        var result = $"List item edited successfully\n";
        result += $"Paragraph index: {paragraphIndex}\n";
        result += $"New text: {text}\n";
        if (level.HasValue)
        {
            result += $"Level: {level.Value}\n";
        }
        result += $"Output: {outputPath}";
        
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
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
        var numberStyle = ArgumentHelper.GetStringNullable(arguments, "numberStyle");
        var indentLevel = ArgumentHelper.GetIntNullable(arguments, "indentLevel");
        var leftIndent = ArgumentHelper.GetDoubleNullable(arguments, "leftIndent");
        var firstLineIndent = ArgumentHelper.GetDoubleNullable(arguments, "firstLineIndent");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");
        }
        
        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null)
        {
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex}");
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
                changes.Add($"Number style: {numberStyle}");
            }
        }
        
        if (indentLevel.HasValue)
        {
            para.ParagraphFormat.LeftIndent = indentLevel.Value * 36;
            changes.Add($"Indent level: {indentLevel.Value}");
        }
        
        if (leftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = leftIndent.Value;
            changes.Add($"Left indent: {leftIndent.Value} points");
        }
        
        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
            changes.Add($"First line indent: {firstLineIndent.Value} points");
        }
        
        doc.Save(outputPath);
        
        var result = $"List format set successfully\n";
        result += $"Paragraph index: {paragraphIndex}\n";
        if (changes.Count > 0)
        {
            result += $"Changes: {string.Join(", ", changes)}\n";
        }
        else
        {
            result += "No change parameters provided\n";
        }
        result += $"Output: {outputPath}";
        
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
        var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var result = new StringBuilder();

        result.AppendLine("=== Document List Format Information ===\n");

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
            
            var para = paragraphs[paragraphIndex.Value];
            AppendListFormatInfo(result, para, paragraphIndex.Value);
        }
        else
        {
            var listParagraphs = paragraphs
                .Where(p => p.ListFormat != null && p.ListFormat.IsListItem)
                .ToList();
            
            result.AppendLine($"Total list paragraphs: {listParagraphs.Count}\n");
            
            if (listParagraphs.Count == 0)
            {
                result.AppendLine("No list paragraphs found");
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
        result.AppendLine($"[Paragraph {paraIndex}]");
        result.AppendLine($"Content preview: {para.GetText().Trim().Substring(0, Math.Min(50, para.GetText().Trim().Length))}...");
        
        if (para.ListFormat != null && para.ListFormat.IsListItem)
        {
            result.AppendLine($"Is list item: Yes");
            result.AppendLine($"List level: {para.ListFormat.ListLevelNumber}");
            
            if (para.ListFormat.List != null)
            {
                result.AppendLine($"List ID: {para.ListFormat.List.ListId}");
            }
            
            if (para.ListFormat.ListLevel != null)
            {
                var level = para.ListFormat.ListLevel;
                result.AppendLine($"List symbol: {level.NumberFormat}");
                result.AppendLine($"Alignment: {level.Alignment}");
                result.AppendLine($"Text position: {level.TextPosition}");
                result.AppendLine($"Number style: {level.NumberStyle}");
            }
        }
        else
        {
            result.AppendLine($"Is list item: No");
            result.AppendLine($"Note: This paragraph is not a list item, cannot get list format information. To convert this paragraph to a list item, use insert_list or set_list_style operation");
        }
    }

    private List<(string text, int level)> ParseItems(JsonNode? itemsNode)
    {
        var items = new List<(string text, int level)>();

        if (itemsNode == null)
        {
            throw new ArgumentException("❌ items parameter cannot be null\n\n" +
                                      $"📝 Please provide an array in the format:\n" +
                                      $"  Simple format: [\"Item 1\", \"Item 2\", \"Item 3\"]\n" +
                                      $"  With level format: [{{\"text\": \"Item 1\", \"level\": 0}}, {{\"text\": \"Sub-item\", \"level\": 1}}]");
        }

        try
        {
            var itemsArray = itemsNode.AsArray();
            if (itemsArray == null)
            {
                var nodeType = itemsNode.GetType().Name;
                var nodeValue = itemsNode.ToString();
                throw new ArgumentException($"❌ items parameter must be an array\n\n" +
                                          $"📋 Current type: {nodeType}\n" +
                                          $"📋 Current value: {nodeValue}\n\n" +
                                          $"📝 Correct format examples:\n" +
                                          $"  Simple format: [\"Item 1\", \"Item 2\", \"Item 3\"]\n" +
                                          $"  With level format: [{{\"text\": \"Item 1\", \"level\": 0}}, {{\"text\": \"Sub-item\", \"level\": 1}}]");
            }
            
            if (itemsArray.Count == 0)
            {
                throw new ArgumentException("❌ items array cannot be empty\n\n" +
                                          $"📝 Please provide at least one item, e.g.: [\"Item 1\"]");
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
                        throw new ArgumentException($"❌ Unable to parse list item as string: {ex.Message}\n\n" +
                                                  $"📋 Item value: {item}\n\n" +
                                                  $"📝 Correct format: string, e.g. \"Item 1\"");
                    }
                }
                else if (item is JsonObject jsonObj)
                {
                    // Object with text and level
                    var text = jsonObj["text"]?.GetValue<string>();
                    if (string.IsNullOrEmpty(text))
                    {
                        var objKeys = string.Join(", ", jsonObj.Select(kvp => $"'{kvp.Key}'"));
                        throw new ArgumentException($"❌ List item object must contain 'text' property\n\n" +
                                                  $"📋 Current object keys: {objKeys}\n\n" +
                                                  $"📝 Correct format: {{\"text\": \"Item text\", \"level\": 0}}");
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
                    throw new ArgumentException($"❌ Invalid list item format\n\n" +
                                              $"📋 Item type: {item.GetType().Name}\n" +
                                              $"📋 Item value: {item}\n\n" +
                                              $"📝 Correct format:\n" +
                                              $"  String: \"Item text\"\n" +
                                              $"  Object: {{\"text\": \"Item text\", \"level\": 0}}");
                }
            }
            
            if (items.Count == 0)
            {
                throw new ArgumentException("❌ No valid list items after parsing\n\n" +
                                          $"📝 Please ensure items array contains at least one valid string or object");
            }
        }
        catch (ArgumentException)
        {
            throw; // Re-throw ArgumentException as-is
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"❌ Error parsing items parameter: {ex.Message}\n\n" +
                                      $"📋 Error type: {ex.GetType().Name}\n\n" +
                                      $"📝 Please ensure items is an array in the format:\n" +
                                      $"  Simple format: [\"Item 1\", \"Item 2\"]\n" +
                                      $"  With level format: [{{\"text\": \"Item 1\", \"level\": 0}}, ...]", ex);
        }

        if (items.Count == 0)
        {
            throw new ArgumentException("Unable to parse any valid list items. Please check items parameter format");
        }

        return items;
    }
}

