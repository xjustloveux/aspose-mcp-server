using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Core;
using static Aspose.Words.ConvertUtil;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for list operations in Word documents
///     Merges: WordAddListTool, WordAddListItemTool, WordDeleteListItemTool, WordEditListItemTool,
///     WordSetListFormatTool, WordGetListFormatTool
/// </summary>
public class WordListTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Manage lists in Word documents. Supports 8 operations: add_list, add_item, delete_item, edit_item, set_format, get_format, restart_numbering, convert_to_list.

Usage examples:
- Add bullet list: word_list(operation='add_list', path='doc.docx', items=['Item 1', 'Item 2', 'Item 3'])
- Add numbered list: word_list(operation='add_list', path='doc.docx', items=['First', 'Second'], listType='number')
- Add list item: word_list(operation='add_item', path='doc.docx', text='New item', styleName='Heading 4')
- Delete list item: word_list(operation='delete_item', path='doc.docx', paragraphIndex=0)
- Edit list item: word_list(operation='edit_item', path='doc.docx', paragraphIndex=0, text='Updated text')
- Get list format: word_list(operation='get_format', path='doc.docx', paragraphIndex=0)
- Restart numbering: word_list(operation='restart_numbering', path='doc.docx', paragraphIndex=2, startAt=1)
- Convert to list: word_list(operation='convert_to_list', path='doc.docx', startParagraphIndex=0, endParagraphIndex=5)

Note: The 'operation' parameter is optional and will be auto-inferred from other parameters. You can also explicitly specify it.";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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
- 'get_format': Get list format (required params: path, paragraphIndex). Note: This operation can only be used on list item paragraphs. If the paragraph is not a list item, it will return a message indicating that the paragraph is not a list item.
- 'restart_numbering': Restart list numbering from 1 at specified paragraph (required params: path, paragraphIndex)
- 'convert_to_list': Convert existing paragraphs to a list (required params: path, startParagraphIndex, endParagraphIndex)",
                @enum = new[]
                {
                    "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format", "restart_numbering",
                    "convert_to_list"
                }
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
Supports two formats:
1. Simple string array: ['Item 1', 'Item 2', 'Item 3']
2. Object array with level (for multi-level/nested lists):
   [{'text': 'Main item', 'level': 0}, {'text': 'Sub-item', 'level': 1}, {'text': 'Sub-sub-item', 'level': 2}]
Level range: 0-8 (0 = top level)",
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
                description = "Custom bullet character (optional, for custom type, e.g., '��', '��', '?')"
            },
            numberFormat = new
            {
                type = "string",
                description =
                    "Number format for numbered lists: arabic, roman, letter (optional, default: arabic, for add_list operation)",
                @enum = new[] { "arabic", "roman", "letter" }
            },
            continuePrevious = new
            {
                type = "boolean",
                description =
                    "If true, continues numbering from the last list in the document instead of starting a new list (optional, default: false, for add_list operation)"
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
                description =
                    "Style name for the list item (required for add_item operation). Example: 'Heading 4'. Use word_get_styles tool to see available styles."
            },
            listLevel = new
            {
                type = "number",
                description = "List level (0-8, optional, for add_item operation)"
            },
            applyStyleIndent = new
            {
                type = "boolean",
                description =
                    "If true, uses the indentation defined in the style (optional, default: true, for add_item operation)"
            },
            // Delete/Edit item parameters
            paragraphIndex = new
            {
                type = "number",
                description =
                    "Paragraph index (0-based, required for delete_item, edit_item, set_format, and get_format operations). Note: For get_format operation, this must be a list item paragraph. If the paragraph is not a list item, the operation will return a message indicating that the paragraph is not a list item."
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
                description =
                    "Indentation level (0-8, optional, for set_format operation). Each level = 36 points (0.5 inch)"
            },
            leftIndent = new
            {
                type = "number",
                description =
                    "Left indent in points (optional, overrides indentLevel if provided, for set_format operation)"
            },
            firstLineIndent = new
            {
                type = "number",
                description =
                    "First line indent in points (optional, negative for hanging indent, for set_format operation)"
            },
            // Restart numbering parameters
            startAt = new
            {
                type = "number",
                description = "Number to restart at (optional, default: 1, for restart_numbering operation)"
            },
            // Convert to list parameters
            startParagraphIndex = new
            {
                type = "number",
                description = "Starting paragraph index (0-based, required for convert_to_list operation)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description = "Ending paragraph index (0-based, inclusive, required for convert_to_list operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        if (arguments == null)
            throw new ArgumentException("? Arguments cannot be null\n\n" +
                                        "?? Usage example: word_list(path='doc.docx', items=['Item 1', 'Item 2'])");

        if (!arguments.ContainsKey("path"))
        {
            var providedKeys = arguments.Select(kvp => kvp.Key).ToList();
            throw new ArgumentException($"? Required parameter 'path' is missing\n\n" +
                                        $"?? Provided parameters: {(providedKeys.Count > 0 ? string.Join(", ", providedKeys.Select(k => $"'{k}'")) : "none")}\n\n" +
                                        $"?? Usage examples:\n" +
                                        $"  word_list(path='doc.docx', items=['Item 1', 'Item 2', 'Item 3'])\n" +
                                        $"  word_list(path='doc.docx', text='New item', styleName='Heading 4')\n" +
                                        $"  word_list(path='doc.docx', paragraphIndex=0)\n\n" +
                                        $"?? Note: 'path' parameter is required for all operations.");
        }

        var pathValue = arguments["path"];
        if (pathValue == null)
            throw new ArgumentException("? Parameter 'path' is null\n\n" +
                                        "?? Usage example: word_list(path='doc.docx', items=['Item 1', 'Item 2'])\n\n" +
                                        "?? Note: 'path' must be a non-null string value.");

        string path;
        try
        {
            path = pathValue.GetValue<string>();
        }
        catch (Exception ex)
        {
            var pathType = pathValue.GetType().Name;
            throw new ArgumentException($"? Parameter 'path' has incorrect type\n\n" +
                                        $"?? Current type: {pathType}\n" +
                                        $"?? Current value: {pathValue}\n\n" +
                                        $"?? Expected: string (e.g., 'doc.docx')\n\n" +
                                        $"?? Error: {ex.Message}");
        }

        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("? Parameter 'path' cannot be empty\n\n" +
                                        "?? Usage example: word_list(path='doc.docx', items=['Item 1', 'Item 2'])\n\n" +
                                        "?? Note: 'path' must be a non-empty string containing the document file path.");

        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);

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
                    // Has text and itemIndex -> edit_item
                    operation = "edit_item";
                else
                    // Has text but no itemIndex -> add_item
                    operation = "add_item";
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
                        // Same path and no text -> get_format (read operation)
                        operation = "get_format";
                    else
                        // Different path or has text -> delete_item
                        operation = "delete_item";
                }
            }
            else
            {
                // Cannot infer operation
                var availableOps = new[]
                    { "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format" };
                throw new ArgumentException(
                    $"? Required parameter 'operation' is missing and cannot be inferred from provided parameters\n\n" +
                    $"?? {providedParamsInfo}\n\n" +
                    $"?? Available operations: {string.Join(", ", availableOps)}\n\n" +
                    $"?? Usage examples:\n" +
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
                    $"?? Tip: If auto-inference fails, explicitly specify the operation parameter");
            }

            // Add the inferred operation to arguments for consistency
            arguments["operation"] = operation;
        }
        else
        {
            operation = ArgumentHelper.GetString(arguments, "operation");

            // Validate operation value
            var validOperations = new[]
            {
                "add_list", "add_item", "delete_item", "edit_item", "set_format", "get_format", "restart_numbering",
                "convert_to_list"
            };
            if (!validOperations.Contains(operation))
                throw new ArgumentException(
                    $"Invalid operation: '{operation}'. Valid operations: {string.Join(", ", validOperations.Select(op => $"'{op}'"))}");
        }

        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation switch
        {
            "add_list" => await AddListAsync(path, outputPath, arguments),
            "add_item" => await AddListItemAsync(path, outputPath, arguments),
            "delete_item" => await DeleteListItemAsync(path, outputPath, arguments),
            "edit_item" => await EditListItemAsync(path, outputPath, arguments),
            "set_format" => await SetListFormatAsync(path, outputPath, arguments),
            "get_format" => await GetListFormatAsync(path, arguments),
            "restart_numbering" => await RestartNumberingAsync(path, outputPath, arguments),
            "convert_to_list" => await ConvertToListAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a list to the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing items array, optional listType, listStyle, outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> AddListAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var items = arguments?["items"];
            if (items == null) throw new ArgumentException("? items parameter is required");

            try
            {
                var parsedItems = ParseItems(items);
                var listType = ArgumentHelper.GetString(arguments, "listType", "bullet");
                var bulletChar = ArgumentHelper.GetString(arguments, "bulletChar", "•");
                var numberFormat = ArgumentHelper.GetString(arguments, "numberFormat", "arabic");
                var continuePrevious = ArgumentHelper.GetBool(arguments, "continuePrevious", false);

                // Open document and create list
                var doc = new Document(path);
                var builder = new DocumentBuilder(doc);
                builder.MoveToDocumentEnd();

                List? list = null;
                var isContinuing = false;

                // Try to continue from previous list if requested
                if (continuePrevious && doc.Lists.Count > 0)
                {
                    // Find the last list in the document
                    var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    for (var i = paragraphs.Count - 1; i >= 0; i--)
                        if (paragraphs[i].ListFormat is { IsListItem: true, List: not null })
                        {
                            list = paragraphs[i].ListFormat.List;
                            isContinuing = true;
                            break;
                        }
                }

                // Create new list if not continuing
                if (list == null)
                {
                    list = doc.Lists.Add(listType == "number"
                        ? ListTemplate.NumberDefault
                        : ListTemplate.BulletDefault);

                    // Configure list format for new lists only
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

                        foreach (var level in list.ListLevels) level.NumberStyle = numStyle;
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

                var result = isContinuing
                    ? "List items added (continuing previous list)\n"
                    : "List added successfully\n";
                if (!isContinuing)
                {
                    result += $"Type: {listType}\n";
                    if (listType == "custom") result += $"Bullet character: {bulletChar}\n";
                    if (listType == "number") result += $"Number format: {numberFormat}\n";
                }
                else
                {
                    result += $"Continued from list ID: {list.ListId}\n";
                }

                result += $"Item count: {parsedItems.Count}\n";
                result += $"Output: {outputPath}";

                return result;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"? Error creating list: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Adds an item to an existing list
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing listIndex, text, optional insertAt, outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> AddListItemAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var text = ArgumentHelper.GetString(arguments, "text");
            var styleName = ArgumentHelper.GetString(arguments, "styleName");
            var listLevel = ArgumentHelper.GetInt(arguments, "listLevel", 0);
            var applyStyleIndent = ArgumentHelper.GetBool(arguments, "applyStyleIndent", true);

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            var style = doc.Styles[styleName];
            if (style == null)
            {
                // Suggest common list-related styles
                var commonStyles = new[]
                {
                    "List Paragraph", "List Bullet", "List Number", "Heading 1", "Heading 2", "Heading 3", "Heading 4"
                };
                var availableCommon = commonStyles.Where(s => doc.Styles[s] != null).Take(3).ToList();
                var suggestions = availableCommon.Count > 0
                    ? $"Common available styles: {string.Join(", ", availableCommon.Select(s => $"'{s}'"))}"
                    : "Use word_get_styles tool to view available styles";
                throw new ArgumentException(
                    $"Style '{styleName}' not found. {suggestions}");
            }

            var para = new Paragraph(doc)
            {
                ParagraphFormat = { StyleName = styleName }
            };

            if (!applyStyleIndent && listLevel > 0) para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * listLevel);

            var run = new Run(doc, text);
            para.AppendChild(run);
            builder.CurrentParagraph.ParentNode.AppendChild(para);

            doc.Save(outputPath);

            var result = "List item added successfully\n";
            result += $"Style: {styleName}\n";
            result += $"Level: {listLevel}\n";

            if (applyStyleIndent)
                result += "Indent: Using style-defined indent (recommended)\n";
            else if (listLevel > 0) result += $"Indent: Manually set ({listLevel * 36} points)\n";

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes an item from a list
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing listIndex, itemIndex, optional outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteListItemAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

            if (paragraphs[paragraphIndex] is not Paragraph paraToDelete)
                throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex}");

            var itemText = paraToDelete.GetText().Trim();
            var itemPreview = itemText.Length > 50 ? itemText.Substring(0, 50) + "..." : itemText;
            var isListItem = paraToDelete.ListFormat.IsListItem;
            var listInfo = isListItem ? " (list item)" : " (regular paragraph)";

            paraToDelete.Remove();
            doc.Save(outputPath);

            var result = $"List item #{paragraphIndex} deleted successfully{listInfo}\n";
            if (!string.IsNullOrEmpty(itemPreview)) result += $"Content preview: {itemPreview}\n";
            result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Edits a list item
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing listIndex, itemIndex, text, optional outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> EditListItemAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var text = ArgumentHelper.GetString(arguments, "text");
            var level = ArgumentHelper.GetIntNullable(arguments, "level");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

            if (paragraphs[paragraphIndex] is not Paragraph para)
                throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex}");

            para.Runs.Clear();
            var run = new Run(doc, text);
            para.AppendChild(run);

            if (level is >= 0 and <= 8) para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * level.Value);

            doc.Save(outputPath);

            var result = "List item edited successfully\n";
            result += $"Paragraph index: {paragraphIndex}\n";
            result += $"New text: {text}\n";
            if (level.HasValue) result += $"Level: {level.Value}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Sets list format properties
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing listIndex, optional listType, listStyle, formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> SetListFormatAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var numberStyle = ArgumentHelper.GetStringNullable(arguments, "numberStyle");
            var indentLevel = ArgumentHelper.GetIntNullable(arguments, "indentLevel");
            var leftIndent = ArgumentHelper.GetDoubleNullable(arguments, "leftIndent");
            var firstLineIndent = ArgumentHelper.GetDoubleNullable(arguments, "firstLineIndent");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paragraphIndex] as Paragraph;
            if (para == null)
                throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex}");

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
                    changes.Add($"Number style: {numberStyle} (affects all items at level {level} in this list)");
                }
            }

            if (indentLevel.HasValue)
            {
                para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * indentLevel.Value);
                changes.Add($"Indent level: {indentLevel.Value} ({InchToPoint(0.5 * indentLevel.Value):F1} points)");
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

            var result = "List format set successfully\n";
            result += $"Paragraph index: {paragraphIndex}\n";
            if (changes.Count > 0)
                result += $"Changes: {string.Join(", ", changes)}\n";
            else
                result += "No change parameters provided\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }


    /// <summary>
    ///     Gets list format information
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing listIndex</param>
    /// <returns>JSON formatted string with list format details</returns>
    private Task<string> GetListFormatAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            // Build list item index mapping for each list
            var listItemIndices = new Dictionary<(int listId, int paraIndex), int>();
            var listCounters = new Dictionary<int, int>();
            foreach (var para in paragraphs)
                if (para.ListFormat is { IsListItem: true, List: not null })
                {
                    var listId = para.ListFormat.List.ListId;
                    listCounters.TryAdd(listId, 0);
                    var paraIdx = paragraphs.IndexOf(para);
                    listItemIndices[(listId, paraIdx)] = listCounters[listId];
                    listCounters[listId]++;
                }

            if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

                var para = paragraphs[paragraphIndex.Value];
                var listInfo = BuildListFormatInfo(para, paragraphIndex.Value, listItemIndices);

                return JsonSerializer.Serialize(listInfo, new JsonSerializerOptions { WriteIndented = true });
            }

            var listParagraphs = paragraphs
                .Where(p => p.ListFormat is { IsListItem: true })
                .ToList();

            if (listParagraphs.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    listParagraphs = Array.Empty<object>(),
                    message = "No list paragraphs found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var listInfos = new List<object>();
            foreach (var para in listParagraphs)
            {
                var paraIndex = paragraphs.IndexOf(para);
                listInfos.Add(BuildListFormatInfo(para, paraIndex, listItemIndices));
            }

            var result = new
            {
                count = listParagraphs.Count,
                listParagraphs = listInfos
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Builds list format information for a paragraph
    /// </summary>
    /// <param name="para">Paragraph to get format info from</param>
    /// <param name="paraIndex">Index of the paragraph in the document</param>
    /// <param name="listItemIndices">Dictionary mapping (listId, paraIndex) to list item index</param>
    /// <returns>Object with list format details</returns>
    private static object BuildListFormatInfo(Paragraph para, int paraIndex,
        Dictionary<(int listId, int paraIndex), int> listItemIndices)
    {
        var previewText = para.ToString(SaveFormat.Text).Trim();
        if (previewText.Length > 50) previewText = previewText[..50] + "...";

        if (para.ListFormat is { IsListItem: true })
        {
            var listInfo = new Dictionary<string, object?>
            {
                ["paragraphIndex"] = paraIndex,
                ["contentPreview"] = previewText,
                ["isListItem"] = true,
                ["listLevel"] = para.ListFormat.ListLevelNumber
            };

            if (para.ListFormat.List != null)
            {
                var listId = para.ListFormat.List.ListId;
                listInfo["listId"] = listId;

                if (listItemIndices.TryGetValue((listId, paraIndex), out var listItemIndex))
                    listInfo["listItemIndex"] = listItemIndex;
            }

            if (para.ListFormat.ListLevel != null)
            {
                var level = para.ListFormat.ListLevel;
                listInfo["listLevelFormat"] = new
                {
                    symbol = level.NumberFormat,
                    alignment = level.Alignment.ToString(),
                    textPosition = level.TextPosition,
                    numberStyle = level.NumberStyle.ToString()
                };
            }

            return listInfo;
        }

        return new
        {
            paragraphIndex = paraIndex,
            contentPreview = previewText,
            isListItem = false,
            note = "This paragraph is not a list item. Use convert_to_list operation to convert it."
        };
    }

    /// <summary>
    ///     Restarts list numbering at specified paragraph
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional startAt</param>
    /// <returns>Success message</returns>
    private Task<string> RestartNumberingAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var startAt = ArgumentHelper.GetInt(arguments, "startAt", 1);

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paragraphIndex];

            if (!para.ListFormat.IsListItem)
                throw new ArgumentException(
                    $"Paragraph at index {paragraphIndex} is not a list item. Use get_format operation to find list item paragraphs.");

            var originalList = para.ListFormat.List;
            if (originalList == null)
                throw new InvalidOperationException("Unable to access list for this paragraph");

            // Create a copy of the list to restart numbering
            var newList = doc.Lists.AddCopy(originalList);
            var level = para.ListFormat.ListLevelNumber;

            // Set the starting number
            newList.ListLevels[level].StartAt = startAt;

            // Apply the new list to this paragraph and all following paragraphs in the same original list
            var applyCount = 0;
            for (var i = paragraphIndex; i < paragraphs.Count; i++)
            {
                var p = paragraphs[i];
                if (p.ListFormat.IsListItem && p.ListFormat.List?.ListId == originalList.ListId)
                {
                    p.ListFormat.List = newList;
                    applyCount++;
                }
                else if (i > paragraphIndex && !p.ListFormat.IsListItem)
                {
                    // Stop when we hit a non-list paragraph after the starting point
                    break;
                }
            }

            doc.Save(outputPath);

            var result = "List numbering restarted successfully\n";
            result += $"Paragraph index: {paragraphIndex}\n";
            result += $"Start at: {startAt}\n";
            result += $"Paragraphs affected: {applyCount}\n";
            result += $"New list ID: {newList.ListId}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Converts existing paragraphs to a list
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing startParagraphIndex, endParagraphIndex, optional listType</param>
    /// <returns>Success message</returns>
    private Task<string> ConvertToListAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startIndex = ArgumentHelper.GetInt(arguments, "startParagraphIndex");
            var endIndex = ArgumentHelper.GetInt(arguments, "endParagraphIndex");
            var listType = ArgumentHelper.GetString(arguments, "listType", "bullet");
            var numberFormat = ArgumentHelper.GetString(arguments, "numberFormat", "arabic");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            if (startIndex < 0 || startIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Start paragraph index {startIndex} is out of range (document has {paragraphs.Count} paragraphs)");

            if (endIndex < 0 || endIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"End paragraph index {endIndex} is out of range (document has {paragraphs.Count} paragraphs)");

            if (startIndex > endIndex)
                throw new ArgumentException(
                    $"Start index ({startIndex}) must be less than or equal to end index ({endIndex})");

            // Create list
            var list = doc.Lists.Add(listType == "number"
                ? ListTemplate.NumberDefault
                : ListTemplate.BulletDefault);

            // Configure number format if needed
            if (listType == "number")
            {
                var numStyle = numberFormat.ToLower() switch
                {
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    _ => NumberStyle.Arabic
                };

                foreach (var level in list.ListLevels) level.NumberStyle = numStyle;
            }

            // Apply list to paragraphs
            var convertedCount = 0;
            var skippedCount = 0;
            for (var i = startIndex; i <= endIndex; i++)
            {
                var para = paragraphs[i];

                // Skip paragraphs that are already list items
                if (para.ListFormat.IsListItem)
                {
                    skippedCount++;
                    continue;
                }

                // Skip empty paragraphs
                var text = para.ToString(SaveFormat.Text).Trim();
                if (string.IsNullOrEmpty(text))
                {
                    skippedCount++;
                    continue;
                }

                para.ListFormat.List = list;
                para.ListFormat.ListLevelNumber = 0;
                convertedCount++;
            }

            doc.Save(outputPath);

            var result = "Paragraphs converted to list successfully\n";
            result += $"Range: paragraph {startIndex} to {endIndex}\n";
            result += $"List type: {listType}\n";
            if (listType == "number") result += $"Number format: {numberFormat}\n";
            result += $"Converted: {convertedCount} paragraphs\n";
            if (skippedCount > 0) result += $"Skipped: {skippedCount} paragraphs (already list items or empty)\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Parses list items from JSON node
    /// </summary>
    /// <param name="itemsNode">JSON node containing list items (string array or object array with text/level)</param>
    /// <returns>List of tuples containing item text and level</returns>
    private List<(string text, int level)> ParseItems(JsonNode? itemsNode)
    {
        var items = new List<(string text, int level)>();

        if (itemsNode == null)
            throw new ArgumentException("? items parameter cannot be null\n\n" +
                                        "?? Please provide an array in the format:\n" +
                                        "  Simple format: [\"Item 1\", \"Item 2\", \"Item 3\"]\n" +
                                        "  With level format: [{\"text\": \"Item 1\", \"level\": 0}, {\"text\": \"Sub-item\", \"level\": 1}]");

        try
        {
            if (itemsNode is not JsonArray itemsArray)
            {
                var nodeType = itemsNode.GetType().Name;
                var nodeValue = itemsNode.ToString();
                throw new ArgumentException($"? items parameter must be an array\n\n" +
                                            $"?? Current type: {nodeType}\n" +
                                            $"?? Current value: {nodeValue}\n\n" +
                                            $"?? Correct format examples:\n" +
                                            $"  Simple format: [\"Item 1\", \"Item 2\", \"Item 3\"]\n" +
                                            $"  With level format: [{{\"text\": \"Item 1\", \"level\": 0}}, {{\"text\": \"Sub-item\", \"level\": 1}}]");
            }

            if (itemsArray.Count == 0)
                throw new ArgumentException("? items array cannot be empty\n\n" +
                                            "?? Please provide at least one item, e.g.: [\"Item 1\"]");

            foreach (var item in itemsArray)
            {
                if (item == null) continue; // Skip null items

                if (item is JsonValue jsonValue)
                {
                    // Simple string item
                    try
                    {
                        var text = jsonValue.GetValue<string>();
                        if (!string.IsNullOrEmpty(text)) items.Add((text, 0));
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException($"? Unable to parse list item as string: {ex.Message}\n\n" +
                                                    $"?? Item value: {item}\n\n" +
                                                    $"?? Correct format: string, e.g. \"Item 1\"");
                    }
                }
                else if (item is JsonObject jsonObj)
                {
                    // Object with text and level
                    var text = jsonObj["text"]?.GetValue<string>();
                    if (string.IsNullOrEmpty(text))
                    {
                        var objKeys = string.Join(", ", jsonObj.Select(kvp => $"'{kvp.Key}'"));
                        throw new ArgumentException($"? List item object must contain 'text' property\n\n" +
                                                    $"?? Current object keys: {objKeys}\n\n" +
                                                    $"?? Correct format: {{\"text\": \"Item text\", \"level\": 0}}");
                    }

                    var level = jsonObj["level"]?.GetValue<int>() ?? 0;
                    if (level < 0 || level > 8) level = Math.Max(0, Math.Min(8, level)); // Clamp to valid range

                    items.Add((text, level));
                }
                else
                {
                    throw new ArgumentException($"? Invalid list item format\n\n" +
                                                $"?? Item type: {item.GetType().Name}\n" +
                                                $"?? Item value: {item}\n\n" +
                                                $"?? Correct format:\n" +
                                                $"  String: \"Item text\"\n" +
                                                $"  Object: {{\"text\": \"Item text\", \"level\": 0}}");
                }
            }

            if (items.Count == 0)
                throw new ArgumentException("? No valid list items after parsing\n\n" +
                                            "?? Please ensure items array contains at least one valid string or object");
        }
        catch (ArgumentException)
        {
            throw; // Re-throw ArgumentException as-is
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"? Error parsing items parameter: {ex.Message}\n\n" +
                                        $"?? Error type: {ex.GetType().Name}\n\n" +
                                        $"?? Please ensure items is an array in the format:\n" +
                                        $"  Simple format: [\"Item 1\", \"Item 2\"]\n" +
                                        $"  With level format: [{{\"text\": \"Item 1\", \"level\": 0}}, ...]", ex);
        }

        if (items.Count == 0)
            throw new ArgumentException("Unable to parse any valid list items. Please check items parameter format");

        return items;
    }
}