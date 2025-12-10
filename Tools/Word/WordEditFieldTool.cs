using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordEditFieldTool : IAsposeTool
{
    public string Description => "Edit field parameters and formatting in a Word document";

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
                description = "Field index to edit (0-based)"
            },
            fieldCode = new
            {
                type = "string",
                description = "New field code (optional, e.g., 'DATE \\@ \"yyyy-MM-dd\"')"
            },
            lockField = new
            {
                type = "boolean",
                description = "Lock the field (optional)"
            },
            unlockField = new
            {
                type = "boolean",
                description = "Unlock the field (optional)"
            },
            updateField = new
            {
                type = "boolean",
                description = "Update the field after editing (default: true)"
            }
        },
        required = new[] { "path", "fieldIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int>() ?? throw new ArgumentException("fieldIndex is required");
        var fieldCode = arguments?["fieldCode"]?.GetValue<string>();
        var lockField = arguments?["lockField"]?.GetValue<bool?>();
        var unlockField = arguments?["unlockField"]?.GetValue<bool?>();
        var updateField = arguments?["updateField"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var fields = doc.Range.Fields.ToList();
        
        if (fieldIndex < 0 || fieldIndex >= fields.Count)
        {
            throw new ArgumentException($"字段索引 {fieldIndex} 超出範圍 (文檔共有 {fields.Count} 個字段)");
        }

        var field = fields[fieldIndex];
        var oldFieldCode = field.GetFieldCode();
        var oldResult = field.Result ?? "";
        var oldLocked = field.IsLocked;
        
        var changes = new List<string>();
        
        // Update field code if provided
        if (!string.IsNullOrEmpty(fieldCode))
        {
            try
            {
                // To change field code, we need to replace the field
                // Get the field start and end nodes
                var fieldStart = field.Start;
                var fieldEnd = field.End;
                
                if (fieldStart != null && fieldEnd != null)
                {
                    // Create a new field with the new code
                    var builder = new DocumentBuilder(doc);
                    builder.MoveTo(fieldStart);
                    
                    // Remove old field separator and result, keep only start and end
                    var currentNode = fieldStart.NextSibling;
                    while (currentNode != null && currentNode != fieldEnd)
                    {
                        var nextNode = currentNode.NextSibling;
                        if (currentNode.NodeType == NodeType.FieldSeparator)
                        {
                            // Keep separator, but we'll replace the code
                        }
                        else if (currentNode.NodeType != NodeType.FieldEnd)
                        {
                            currentNode.Remove();
                        }
                        currentNode = nextNode;
                    }
                    
                    // Insert new field code
                    builder.MoveTo(fieldStart);
                    builder.Write(fieldCode);
                    
                    changes.Add($"字段代碼已更新: {oldFieldCode} -> {fieldCode}");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法更新字段代碼: {ex.Message}", ex);
            }
        }
        
        // Lock/unlock field
        if (lockField.HasValue && lockField.Value)
        {
            field.IsLocked = true;
            changes.Add("字段已鎖定");
        }
        else if (unlockField.HasValue && unlockField.Value)
        {
            field.IsLocked = false;
            changes.Add("字段已解鎖");
        }
        
        // Update field if requested
        if (updateField)
        {
            field.Update();
            doc.UpdateFields();
        }
        
        doc.Save(outputPath);
        
        var result = $"成功編輯字段 #{fieldIndex}\n";
        result += $"原字段代碼: {oldFieldCode}\n";
        if (!string.IsNullOrEmpty(fieldCode))
        {
            result += $"新字段代碼: {fieldCode}\n";
        }
        result += $"原結果: {oldResult}\n";
        result += $"原鎖定狀態: {(oldLocked ? "已鎖定" : "未鎖定")}\n";
        if (changes.Count > 0)
        {
            result += $"變更: {string.Join(", ", changes)}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

