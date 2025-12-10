using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordEditHyperlinkTool : IAsposeTool
{
    public string Description => "Edit an existing hyperlink in Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            hyperlinkIndex = new
            {
                type = "number",
                description = "Hyperlink index (0-based, from word_get_hyperlinks)"
            },
            url = new
            {
                type = "string",
                description = "New URL or target address"
            },
            displayText = new
            {
                type = "string",
                description = "New display text (optional)"
            },
            tooltip = new
            {
                type = "string",
                description = "New tooltip text (optional)"
            }
        },
        required = new[] { "path", "hyperlinkIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required");
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
            // Update the field result (display text)
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
}

