using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordGetHyperlinksTool : IAsposeTool
{
    public string Description => "Get all hyperlinks from Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

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
                    // Get display text from the field result
                    displayText = field.Result ?? "";
                    
                    // Get address
                    address = hyperlinkField.Address ?? "";
                    
                    // Get tooltip
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

