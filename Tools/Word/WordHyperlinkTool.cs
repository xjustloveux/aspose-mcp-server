using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Word hyperlinks (add, edit, delete, get)
/// Merges: WordAddHyperlinkTool, WordEditHyperlinkTool, WordDeleteHyperlinkTool, WordGetHyperlinksTool
/// </summary>
public class WordHyperlinkTool : IAsposeTool
{
    public string Description => @"Manage Word hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: word_hyperlink(operation='add', path='doc.docx', text='Link', url='https://example.com', paragraphIndex=0)
- Edit hyperlink: word_hyperlink(operation='edit', path='doc.docx', hyperlinkIndex=0, url='https://newurl.com')
- Delete hyperlink: word_hyperlink(operation='delete', path='doc.docx', hyperlinkIndex=0)
- Get hyperlinks: word_hyperlink(operation='get', path='doc.docx')";

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
                description = "URL or target address (required for add operation, optional for edit operation)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert hyperlink after (0-based, optional, for add operation). When specified, creates a NEW paragraph after the specified paragraph (does not insert into existing paragraph). Valid range: 0 to (total paragraphs - 1), or -1 for document start."
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(arguments, path),
            "edit" => await EditHyperlinkAsync(arguments, path),
            "delete" => await DeleteHyperlinkAsync(arguments, path),
            "get" => await GetHyperlinksAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds a hyperlink to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing text, address, optional displayText, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddHyperlinkAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var text = ArgumentHelper.GetString(arguments, "text");
        var url = ArgumentHelper.GetString(arguments, "url");
        var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
        var tooltip = ArgumentHelper.GetStringNullable(arguments, "tooltip");

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
                    var firstPara = paragraphs[0] as Paragraph;
                    if (firstPara != null)
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
                var targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                if (targetPara != null)
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
                        throw new InvalidOperationException($"Unable to find parent node of paragraph at index {paragraphIndex.Value}");
                    }
                }
                else
                {
                    throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");
                }
            }
            else
            {
                throw new ArgumentException($"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
        }
        else
        {
            // Default: Move to end of document
            builder.MoveToDocumentEnd();
        }
        
        // Insert hyperlink
        builder.InsertHyperlink(text, url, false);
        
        // Set tooltip if provided
        if (!string.IsNullOrEmpty(tooltip))
        {
            // Get the last inserted field (should be the hyperlink field)
            var fields = doc.Range.Fields;
            if (fields.Count > 0)
            {
                var lastField = fields[fields.Count - 1];
                if (lastField is FieldHyperlink hyperlinkField)
                {
                    hyperlinkField.ScreenTip = tooltip;
                }
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"Hyperlink added successfully\n";
        result += $"Display text: {text}\n";
        result += $"URL: {url}\n";
        if (!string.IsNullOrEmpty(tooltip))
        {
            result += $"Tooltip: {tooltip}\n";
        }
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                result += "Insert position: beginning of document\n";
            }
            else
            {
                result += $"Insert position: after paragraph #{paragraphIndex.Value}\n";
            }
        }
        else
        {
            result += "Insert position: end of document\n";
        }
        result += $"Output: {outputPath}";
        
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Edits an existing hyperlink
    /// </summary>
    /// <param name="arguments">JSON arguments containing hyperlinkIndex, optional text, address, displayText, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditHyperlinkAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var hyperlinkIndex = ArgumentHelper.GetInt(arguments, "hyperlinkIndex");
        var url = ArgumentHelper.GetStringNullable(arguments, "url");
        var displayText = ArgumentHelper.GetStringNullable(arguments, "displayText");
        var tooltip = ArgumentHelper.GetStringNullable(arguments, "tooltip");

        var doc = new Document(path);
        
        // Get all hyperlink fields
        var hyperlinkFields = new List<FieldHyperlink>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldHyperlink linkField)
            {
                hyperlinkFields.Add(linkField);
            }
        }
        
        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinkFields.Count)
        {
            var availableInfo = hyperlinkFields.Count > 0 
                ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                : " (document has no hyperlinks)";
            throw new ArgumentException($"Hyperlink index {hyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
        }
        
        var hyperlinkField = hyperlinkFields[hyperlinkIndex];
        var changes = new List<string>();
        
        // Update URL if provided
        if (!string.IsNullOrEmpty(url))
        {
            hyperlinkField.Address = url;
            changes.Add($"URL: {url}");
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
    /// Deletes a hyperlink from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing hyperlinkIndex, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteHyperlinkAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var hyperlinkIndex = ArgumentHelper.GetInt(arguments, "hyperlinkIndex");

        var doc = new Document(path);
        
        // Get all hyperlink fields
        var hyperlinkFields = new List<FieldHyperlink>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldHyperlink linkField)
            {
                hyperlinkFields.Add(linkField);
            }
        }
        
        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinkFields.Count)
        {
            var availableInfo = hyperlinkFields.Count > 0 
                ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                : " (document has no hyperlinks)";
            throw new ArgumentException($"Hyperlink index {hyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
        }
        
        var hyperlinkField = hyperlinkFields[hyperlinkIndex];
        
        // Get hyperlink info before deletion
        string displayText = hyperlinkField.Result ?? "";
        string address = hyperlinkField.Address ?? "";
        
        // Delete the hyperlink field
        try
        {
            var fieldStart = hyperlinkField.Start;
            var fieldEnd = hyperlinkField.End;
            
            fieldStart.Remove();
            if (fieldEnd != null)
            {
                fieldEnd.Remove();
            }
        }
        catch
        {
            try
            {
                hyperlinkField.Remove();
            }
            catch
            {
                throw new InvalidOperationException("Unable to delete hyperlink, please check document structure");
            }
        }
        
        doc.Save(outputPath);
        
        // Count remaining hyperlinks
        int remainingCount = 0;
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldHyperlink)
            {
                remainingCount++;
            }
        }
        
        var result = $"Hyperlink #{hyperlinkIndex} deleted successfully\n";
        result += $"Display text: {displayText}\n";
        result += $"Address: {address}\n";
        result += $"Remaining hyperlinks in document: {remainingCount}\n";
        result += $"Output: {outputPath}";
        
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Gets all hyperlinks from the document
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all hyperlinks</returns>
    private async Task<string> GetHyperlinksAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        
        // Get all hyperlink fields
        var hyperlinks = new List<(int index, string displayText, string address, string tooltip)>();
        int index = 0;
        
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldHyperlink hyperlinkField)
            {
                string displayText = "";
                string address = "";
                string tooltip = "";
                
                try
                {
                    displayText = field.Result ?? "";
                    address = hyperlinkField.Address ?? "";
                    tooltip = hyperlinkField.ScreenTip ?? "";
                }
                catch
                {
                    // Ignore errors
                }
                
                hyperlinks.Add((index, displayText, address, tooltip));
                index++;
            }
        }
        
        if (hyperlinks.Count == 0)
        {
            return await Task.FromResult("No hyperlinks found in document");
        }
        
        var result = new System.Text.StringBuilder();
        result.AppendLine($"Found {hyperlinks.Count} hyperlinks:\n");
        
        for (int i = 0; i < hyperlinks.Count; i++)
        {
            var (idx, displayText, address, tooltip) = hyperlinks[i];
            result.AppendLine($"Hyperlink #{idx}:");
            result.AppendLine($"  Display text: {displayText}");
            result.AppendLine($"  Address: {address}");
            if (!string.IsNullOrEmpty(tooltip))
            {
                result.AppendLine($"  Tooltip: {tooltip}");
            }
            result.AppendLine();
        }
        
        return await Task.FromResult(result.ToString().TrimEnd());
    }
}

