using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word hyperlinks (add, edit, delete, get)
///     Merges: WordAddHyperlinkTool, WordEditHyperlinkTool, WordDeleteHyperlinkTool, WordGetHyperlinksTool
/// </summary>
public class WordHyperlinkTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Word hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: word_hyperlink(operation='add', path='doc.docx', text='Link', url='https://example.com', paragraphIndex=0)
- Edit hyperlink: word_hyperlink(operation='edit', path='doc.docx', hyperlinkIndex=0, url='https://newurl.com')
- Delete hyperlink: word_hyperlink(operation='delete', path='doc.docx', hyperlinkIndex=0)
- Get hyperlinks: word_hyperlink(operation='get', path='doc.docx')";

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
- 'add': Add a hyperlink (required params: path, text, url)
- 'edit': Edit a hyperlink (required params: path, hyperlinkIndex, url)
- 'delete': Delete a hyperlink (required params: path, hyperlinkIndex)
- 'get': Get all hyperlinks (required params: path)",
                @enum = new[] { "add", "edit", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for add/edit/delete operations)"
            },
            text = new
            {
                type = "string",
                description = "Display text for the hyperlink (required for add operation)"
            },
            url = new
            {
                type = "string",
                description =
                    "URL or target address (required for add operation unless subAddress is provided, optional for edit operation)"
            },
            subAddress = new
            {
                type = "string",
                description =
                    "Internal bookmark name for document navigation (e.g., '_Toc123456'). Use with empty url for internal links. (optional, for add/edit operations)"
            },
            paragraphIndex = new
            {
                type = "number",
                description =
                    "Paragraph index to insert hyperlink after (0-based, optional, for add operation). When specified, creates a NEW paragraph after the specified paragraph (does not insert into existing paragraph). Valid range: 0 to (total paragraphs - 1), or -1 for document start."
            },
            tooltip = new
            {
                type = "string",
                description = "Tooltip text (optional, for add/edit operations)"
            },
            hyperlinkIndex = new
            {
                type = "number",
                description = "Hyperlink index (0-based, required for edit/delete operations)"
            },
            displayText = new
            {
                type = "string",
                description = "New display text (optional, for edit operation)"
            },
            keepText = new
            {
                type = "boolean",
                description =
                    "Keep display text when deleting hyperlink (unlink instead of remove, optional, default: false, for delete operation)"
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
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(path, outputPath, arguments),
            "edit" => await EditHyperlinkAsync(path, outputPath, arguments),
            "delete" => await DeleteHyperlinkAsync(path, outputPath, arguments),
            "get" => await GetHyperlinksAsync(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, url, optional subAddress, paragraphIndex, tooltip</param>
    /// <returns>Success message</returns>
    private Task<string> AddHyperlinkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var text = ArgumentHelper.GetString(arguments, "text");
            var url = ArgumentHelper.GetStringNullable(arguments, "url");
            var subAddress = ArgumentHelper.GetStringNullable(arguments, "subAddress");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var tooltip = ArgumentHelper.GetStringNullable(arguments, "tooltip");

            // Validate: either url or subAddress must be provided
            if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(subAddress))
                throw new ArgumentException("Either 'url' or 'subAddress' must be provided for add operation");

            // Validate URL format if provided
            if (!string.IsNullOrEmpty(url))
                ValidateUrlFormat(url);

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            // Determine insertion position
            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    // Insert at the beginning - create new paragraph
                    if (paragraphs.Count > 0)
                    {
                        if (paragraphs[0] is Paragraph firstPara)
                        {
                            // Insert new paragraph before the first paragraph
                            var newPara = new Paragraph(doc);
                            doc.FirstSection.Body.InsertBefore(newPara, firstPara);
                            builder.MoveTo(newPara);
                        }
                        else
                        {
                            builder.MoveToDocumentStart();
                        }
                    }
                    else
                    {
                        builder.MoveToDocumentStart();
                    }
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    // Insert after the specified paragraph - create new paragraph
                    if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                    {
                        // Insert new paragraph after the target paragraph
                        var newPara = new Paragraph(doc);
                        var parentNode = targetPara.ParentNode;
                        if (parentNode != null)
                        {
                            parentNode.InsertAfter(newPara, targetPara);
                            builder.MoveTo(newPara);
                        }
                        else
                        {
                            throw new InvalidOperationException(
                                $"Unable to find parent node of paragraph at index {paragraphIndex.Value}");
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Unable to find paragraph at index {paragraphIndex.Value}");
                    }
                }
                else
                {
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
                }
            }
            else
            {
                // Default: Move to end of document
                builder.MoveToDocumentEnd();
            }

            // Insert hyperlink
            if (!string.IsNullOrEmpty(subAddress))
                // Internal bookmark link
                builder.InsertHyperlink(text, subAddress, true);
            else
                // External URL link
                builder.InsertHyperlink(text, url!, false);

            // Set tooltip and subAddress if provided
            var fields = doc.Range.Fields;
            if (fields.Count > 0)
            {
                var lastField = fields[^1];
                if (lastField is FieldHyperlink hyperlinkField)
                {
                    if (!string.IsNullOrEmpty(tooltip))
                        hyperlinkField.ScreenTip = tooltip;
                    // Set both Address and SubAddress for combined links
                    if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(subAddress))
                    {
                        hyperlinkField.Address = url;
                        hyperlinkField.SubAddress = subAddress;
                    }
                }
            }

            doc.Save(outputPath);

            var result = "Hyperlink added successfully\n";
            result += $"Display text: {text}\n";
            if (!string.IsNullOrEmpty(url)) result += $"URL: {url}\n";
            if (!string.IsNullOrEmpty(subAddress)) result += $"SubAddress (bookmark): {subAddress}\n";
            if (!string.IsNullOrEmpty(tooltip)) result += $"Tooltip: {tooltip}\n";
            if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value == -1)
                    result += "Insert position: beginning of document\n";
                else
                    result += $"Insert position: after paragraph #{paragraphIndex.Value}\n";
            }
            else
            {
                result += "Insert position: end of document\n";
            }

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Edits an existing hyperlink
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing hyperlinkIndex, optional url, subAddress, displayText, tooltip</param>
    /// <returns>Success message</returns>
    private Task<string> EditHyperlinkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var hyperlinkIndex = ArgumentHelper.GetInt(arguments, "hyperlinkIndex");
            var url = ArgumentHelper.GetStringNullable(arguments, "url");
            var subAddress = ArgumentHelper.GetStringNullable(arguments, "subAddress");
            var displayText = ArgumentHelper.GetStringNullable(arguments, "displayText");
            var tooltip = ArgumentHelper.GetStringNullable(arguments, "tooltip");

            var doc = new Document(path);
            var hyperlinkFields = GetAllHyperlinks(doc);

            if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinkFields.Count)
            {
                var availableInfo = hyperlinkFields.Count > 0
                    ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                    : " (document has no hyperlinks)";
                throw new ArgumentException(
                    $"Hyperlink index {hyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
            }

            var hyperlinkField = hyperlinkFields[hyperlinkIndex];
            var changes = new List<string>();

            // Update URL if provided
            if (!string.IsNullOrEmpty(url))
            {
                ValidateUrlFormat(url);
                hyperlinkField.Address = url;
                changes.Add($"URL: {url}");
            }

            // Update subAddress if provided
            if (!string.IsNullOrEmpty(subAddress))
            {
                hyperlinkField.SubAddress = subAddress;
                changes.Add($"SubAddress: {subAddress}");
            }

            // Update display text if provided
            if (!string.IsNullOrEmpty(displayText))
            {
                hyperlinkField.Result = displayText;
                changes.Add($"Display text: {displayText}");
            }

            // Update tooltip if provided
            if (!string.IsNullOrEmpty(tooltip))
            {
                hyperlinkField.ScreenTip = tooltip;
                changes.Add($"Tooltip: {tooltip}");
            }

            // Update the field
            hyperlinkField.Update();

            doc.Save(outputPath);

            var result = $"Hyperlink #{hyperlinkIndex} edited successfully\n";
            if (changes.Count > 0)
                result += $"Changes: {string.Join(", ", changes)}\n";
            else
                result += "No change parameters provided\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes a hyperlink from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing hyperlinkIndex, optional keepText</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteHyperlinkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var hyperlinkIndex = ArgumentHelper.GetInt(arguments, "hyperlinkIndex");
            var keepText = ArgumentHelper.GetBool(arguments, "keepText", false);

            var doc = new Document(path);
            var hyperlinkFields = GetAllHyperlinks(doc);

            if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinkFields.Count)
            {
                var availableInfo = hyperlinkFields.Count > 0
                    ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                    : " (document has no hyperlinks)";
                throw new ArgumentException(
                    $"Hyperlink index {hyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
            }

            var hyperlinkField = hyperlinkFields[hyperlinkIndex];

            // Get hyperlink info before deletion
            var displayText = hyperlinkField.Result ?? "";
            var address = hyperlinkField.Address ?? "";

            // Delete the hyperlink field
            if (keepText)
                // Unlink: remove hyperlink but keep display text
                hyperlinkField.Unlink();
            else
                // Remove: delete hyperlink and its content
                hyperlinkField.Remove();

            doc.Save(outputPath);

            var remainingCount = GetAllHyperlinks(doc).Count;

            var result = $"Hyperlink #{hyperlinkIndex} deleted successfully\n";
            result += $"Display text: {displayText}\n";
            result += $"Address: {address}\n";
            result += $"Keep text: {(keepText ? "Yes (unlinked)" : "No (removed)")}\n";
            result += $"Remaining hyperlinks in document: {remainingCount}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Gets all hyperlink fields from the document
    /// </summary>
    /// <param name="doc">The Word document</param>
    /// <returns>List of FieldHyperlink objects</returns>
    private static List<FieldHyperlink> GetAllHyperlinks(Document doc)
    {
        return doc.Range.Fields.OfType<FieldHyperlink>().ToList();
    }

    /// <summary>
    ///     Validates URL format to prevent invalid field commands
    /// </summary>
    /// <param name="url">The URL to validate</param>
    private static void ValidateUrlFormat(string url)
    {
        var validPrefixes = new[] { "http://", "https://", "mailto:", "ftp://", "file://", "#" };
        if (!validPrefixes.Any(prefix => url.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException(
                $"Invalid URL format: '{url}'. URL must start with http://, https://, mailto:, ftp://, file://, or # (for internal links)");
    }

    /// <summary>
    ///     Gets all hyperlinks from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <returns>JSON formatted string with all hyperlinks</returns>
    private Task<string> GetHyperlinksAsync(string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var hyperlinkFields = GetAllHyperlinks(doc);

            if (hyperlinkFields.Count == 0)
                return JsonSerializer.Serialize(new
                    { count = 0, hyperlinks = Array.Empty<object>(), message = "No hyperlinks found in document" });

            // Build paragraph lookup for finding hyperlink positions
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            var hyperlinkList = new List<object>();
            for (var index = 0; index < hyperlinkFields.Count; index++)
            {
                var hyperlinkField = hyperlinkFields[index];
                var displayText = "";
                var address = "";
                var subAddress = "";
                var tooltip = "";
                int? paragraphIndex = null;

                try
                {
                    displayText = hyperlinkField.Result ?? "";
                    address = hyperlinkField.Address ?? "";
                    subAddress = hyperlinkField.SubAddress ?? "";
                    tooltip = hyperlinkField.ScreenTip ?? "";

                    // Find paragraph index
                    var fieldStart = hyperlinkField.Start;
                    if (fieldStart?.ParentNode is Paragraph para)
                    {
                        paragraphIndex = paragraphs.IndexOf(para);
                        if (paragraphIndex == -1) paragraphIndex = null;
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[WARN] Error reading hyperlink properties: {ex.Message}");
                }

                hyperlinkList.Add(new
                {
                    index,
                    displayText,
                    address,
                    subAddress = string.IsNullOrEmpty(subAddress) ? null : subAddress,
                    tooltip = string.IsNullOrEmpty(tooltip) ? null : tooltip,
                    paragraphIndex
                });
            }

            var result = new
            {
                count = hyperlinkFields.Count,
                hyperlinks = hyperlinkList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}