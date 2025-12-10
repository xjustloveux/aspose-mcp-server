using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordDeleteHyperlinkTool : IAsposeTool
{
    public string Description => "Delete a specific hyperlink from Word document";

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
            }
        },
        required = new[] { "path", "hyperlinkIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required");

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
        // Remove the field but keep the text
        var fieldStart = hyperlinkField.Start;
        var fieldEnd = hyperlinkField.End;
        
        // Get the parent node
        var parentNode = fieldStart.ParentNode;
        
        // Remove the field code but keep the result text
        // Actually, we should remove the entire field including its result
        // The simplest way is to remove the field start and end, keeping only the result
        try
        {
            // Remove field start and end, but keep the result text
            var runsToKeep = new List<Run>();
            foreach (Node node in parentNode.GetChildNodes(NodeType.Run, true))
            {
                var run = node as Run;
                if (run != null && run.Text.Contains(displayText))
                {
                    // Keep runs that contain the display text
                    runsToKeep.Add(run);
                }
            }
            
            // Remove the field
            fieldStart.Remove();
            if (fieldEnd != null)
            {
                fieldEnd.Remove();
            }
        }
        catch
        {
            // If removal fails, try a simpler approach: just remove the field
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
}

