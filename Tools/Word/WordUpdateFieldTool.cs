using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordUpdateFieldTool : IAsposeTool
{
    public string Description => "Update one or all fields in a Word document";

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
                description = "Field index to update (0-based). If not provided, updates all fields."
            },
            fieldType = new
            {
                type = "string",
                description = "Update only fields of this type (optional, e.g., 'PAGE', 'DATE'). If not provided, updates all fields."
            },
            updateAll = new
            {
                type = "boolean",
                description = "Update all fields in the document (default: false if fieldIndex is provided, true otherwise)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int?>();
        var fieldTypeFilter = arguments?["fieldType"]?.GetValue<string>();
        var updateAll = arguments?["updateAll"]?.GetValue<bool>() ?? (!fieldIndex.HasValue);

        var doc = new Document(path);
        
        var fields = doc.Range.Fields.ToList();
        var updatedCount = 0;
        var errors = new List<string>();
        
        if (fieldIndex.HasValue)
        {
            // Update specific field
            if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            {
                throw new ArgumentException($"功能變數索引 {fieldIndex.Value} 超出範圍 (文檔共有 {fields.Count} 個功能變數)");
            }
            
            var field = fields[fieldIndex.Value];
            try
            {
                var oldResult = field.Result ?? "";
                field.Update();
                var newResult = field.Result ?? "";
                updatedCount = 1;
                
                doc.Save(outputPath);
                
                var result = $"成功更新功能變數 #{fieldIndex.Value}\n";
                result += $"類型: {field.Type}\n";
                result += $"代碼: {field.GetFieldCode()}\n";
                result += $"舊結果: {oldResult}\n";
                result += $"新結果: {newResult}\n";
                result += $"輸出: {outputPath}";
                
                return await Task.FromResult(result);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法更新功能變數 #{fieldIndex.Value}: {ex.Message}", ex);
            }
        }
        else
        {
            // Update all fields or fields of specific type
            foreach (var field in fields)
            {
                // Filter by type if specified
                if (!string.IsNullOrEmpty(fieldTypeFilter))
                {
                    var filterType = fieldTypeFilter.ToUpper();
                    var fieldTypeName = field.Type.ToString().ToUpper();
                    if (!fieldTypeName.Contains(filterType))
                    {
                        continue;
                    }
                }
                
                try
                {
                    field.Update();
                    updatedCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"功能變數 {field.Type} (索引 {fields.IndexOf(field)}): {ex.Message}");
                }
            }
            
            doc.Save(outputPath);
            
            var result = $"成功更新 {updatedCount} 個功能變數\n";
            if (!string.IsNullOrEmpty(fieldTypeFilter))
            {
                result += $"過濾類型: {fieldTypeFilter}\n";
            }
            
            if (errors.Count > 0)
            {
                result += $"\n更新失敗的功能變數 ({errors.Count} 個):\n";
                foreach (var error in errors)
                {
                    result += $"  - {error}\n";
                }
            }
            
            result += $"輸出: {outputPath}";
            
            return await Task.FromResult(result);
        }
    }
}

