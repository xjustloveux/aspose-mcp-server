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
    public string Description => "Manage Word hyperlinks: add, edit, delete, or get all";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'edit', 'delete', 'get'",
                @enum = new[] { "add", "edit", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
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
                description = "Paragraph index to insert hyperlink after (0-based, optional, for add operation)"
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
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(arguments, path),
            "edit" => await EditHyperlinkAsync(arguments, path),
            "delete" => await DeleteHyperlinkAsync(arguments, path),
            "get" => await GetHyperlinksAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddHyperlinkAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required for add operation");
        var url = arguments?["url"]?.GetValue<string>() ?? throw new ArgumentException("url is required for add operation");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var tooltip = arguments?["tooltip"]?.GetValue<string>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                // Insert at the beginning
                if (paragraphs.Count > 0)
                {
                    var firstPara = paragraphs[0] as Paragraph;
                    if (firstPara != null)
                    {
                        builder.MoveTo(firstPara);
                    }
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                // Insert after the specified paragraph
                var targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                if (targetPara != null)
                {
                    builder.MoveTo(targetPara);
                }
                else
                {
                    throw new InvalidOperationException($"無法找到索引 {paragraphIndex.Value} 的段落");
                }
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
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
        
        var result = $"成功添加超連結\n";
        result += $"顯示文字: {text}\n";
        result += $"URL: {url}\n";
        if (!string.IsNullOrEmpty(tooltip))
        {
            result += $"提示文字: {tooltip}\n";
        }
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                result += "插入位置: 文檔開頭\n";
            }
            else
            {
                result += $"插入位置: 段落 #{paragraphIndex.Value} 之後\n";
            }
        }
        else
        {
            result += "插入位置: 文檔末尾\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> EditHyperlinkAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required for edit operation");
        var url = arguments?["url"]?.GetValue<string>();
        var displayText = arguments?["displayText"]?.GetValue<string>();
        var tooltip = arguments?["tooltip"]?.GetValue<string>();

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
            throw new ArgumentException($"超連結索引 {hyperlinkIndex} 超出範圍 (文檔共有 {hyperlinkFields.Count} 個超連結)");
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
            changes.Add($"顯示文字: {displayText}");
        }
        
        // Update tooltip if provided
        if (!string.IsNullOrEmpty(tooltip))
        {
            hyperlinkField.ScreenTip = tooltip;
            changes.Add($"提示文字: {tooltip}");
        }
        
        // Update the field
        hyperlinkField.Update();
        
        doc.Save(outputPath);
        
        var result = $"成功編輯超連結 #{hyperlinkIndex}\n";
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

    private async Task<string> DeleteHyperlinkAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required for delete operation");

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
            throw new ArgumentException($"超連結索引 {hyperlinkIndex} 超出範圍 (文檔共有 {hyperlinkFields.Count} 個超連結)");
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
                throw new InvalidOperationException("無法刪除超連結，請檢查文檔結構");
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
        
        var result = $"成功刪除超連結 #{hyperlinkIndex}\n";
        result += $"顯示文字: {displayText}\n";
        result += $"地址: {address}\n";
        result += $"文檔剩餘超連結數: {remainingCount}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

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
            return await Task.FromResult("文檔中沒有找到超連結");
        }
        
        var result = new System.Text.StringBuilder();
        result.AppendLine($"找到 {hyperlinks.Count} 個超連結：\n");
        
        for (int i = 0; i < hyperlinks.Count; i++)
        {
            var (idx, displayText, address, tooltip) = hyperlinks[i];
            result.AppendLine($"超連結 #{idx}:");
            result.AppendLine($"  顯示文字: {displayText}");
            result.AppendLine($"  地址: {address}");
            if (!string.IsNullOrEmpty(tooltip))
            {
                result.AppendLine($"  提示文字: {tooltip}");
            }
            result.AppendLine();
        }
        
        return await Task.FromResult(result.ToString().TrimEnd());
    }
}

