using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordDeleteFieldTool : IAsposeTool
{
    public string Description => "Delete a field from a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            fieldIndex = new
            {
                type = "number",
                description = "Field index to delete (0-based)"
            },
            keepResult = new
            {
                type = "boolean",
                description = "Keep the field result text after deletion (default: false, removes entire field)"
            }
        },
        required = new[] { "path", "fieldIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int>() ?? throw new ArgumentException("fieldIndex is required");
        var keepResult = arguments?["keepResult"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        var fields = doc.Range.Fields.ToList();
        
        if (fieldIndex < 0 || fieldIndex >= fields.Count)
        {
            throw new ArgumentException($"功能變數索引 {fieldIndex} 超出範圍 (文檔共有 {fields.Count} 個功能變數)");
        }
        
        var field = fields[fieldIndex];
        var fieldType = field.Type.ToString();
        var fieldCode = field.GetFieldCode();
        var fieldResult = field.Result ?? "";
        
        try
        {
            if (keepResult)
            {
                // Keep the result text, remove only the field code
                var fieldStart = field.Start;
                var fieldEnd = field.End;
                
                if (fieldStart != null && fieldEnd != null)
                {
                    // Remove field start and end, keep the result
                    var parentNode = fieldStart.ParentNode;
                    if (parentNode != null)
                    {
                        // Find and remove field start and separator
                        var nodesToRemove = new List<Node>();
                        var currentNode = (Node)fieldStart;
                        var endNode = (Node)fieldEnd;
                        
                        while (currentNode != null && currentNode != endNode.NextSibling)
                        {
                            if (currentNode.NodeType == NodeType.FieldStart || 
                                currentNode.NodeType == NodeType.FieldSeparator)
                            {
                                nodesToRemove.Add(currentNode);
                            }
                            currentNode = currentNode.NextSibling;
                            
                            if (currentNode == endNode)
                            {
                                break;
                            }
                        }
                        
                        foreach (var node in nodesToRemove)
                        {
                            node.Remove();
                        }
                        
                        // Remove field end
                        endNode.Remove();
                    }
                }
            }
            else
            {
                // Remove entire field including result
                var fieldStart = field.Start;
                var fieldEnd = field.End;
                
                if (fieldStart != null && fieldEnd != null)
                {
                    var nodesToRemove = new List<Node>();
                    var currentNode = (Node)fieldStart;
                    var endNode = (Node)fieldEnd;
                    
                    while (currentNode != null)
                    {
                        var nextNode = currentNode.NextSibling;
                        nodesToRemove.Add(currentNode);
                        
                        if (currentNode == endNode)
                        {
                            break;
                        }
                        
                        currentNode = nextNode;
                    }
                    
                    foreach (var node in nodesToRemove)
                    {
                        node.Remove();
                    }
                }
            }
            
            doc.Save(outputPath);
            
            // Count remaining fields
            var remainingFields = doc.Range.Fields.Count;
            
            var result = $"成功刪除功能變數 #{fieldIndex}\n";
            result += $"類型: {fieldType}\n";
            result += $"代碼: {fieldCode}\n";
            if (!string.IsNullOrEmpty(fieldResult))
            {
                result += $"結果: {fieldResult}\n";
            }
            result += $"保留結果文字: {(keepResult ? "是" : "否")}\n";
            result += $"文檔剩餘功能變數數: {remainingFields}\n";
            result += $"輸出: {outputPath}";
            
            return await Task.FromResult(result);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"無法刪除功能變數: {ex.Message}", ex);
        }
    }
}

