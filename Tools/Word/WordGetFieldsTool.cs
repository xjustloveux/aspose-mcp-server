using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordGetFieldsTool : IAsposeTool
{
    public string Description => "Get all fields in a Word document with their types, codes, and results";

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
            fieldType = new
            {
                type = "string",
                description = "Filter by field type (optional, e.g., 'PAGE', 'DATE', 'HYPERLINK'). Leave empty to get all fields."
            },
            includeCode = new
            {
                type = "boolean",
                description = "Include field code in results (default: true)"
            },
            includeResult = new
            {
                type = "boolean",
                description = "Include field result in results (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fieldTypeFilter = arguments?["fieldType"]?.GetValue<string>();
        var includeCode = arguments?["includeCode"]?.GetValue<bool>() ?? true;
        var includeResult = arguments?["includeResult"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        
        var fields = new List<FieldInfo>();
        var fieldIndex = 0;
        
        foreach (Field field in doc.Range.Fields)
        {
            // Filter by type if specified
            if (!string.IsNullOrEmpty(fieldTypeFilter))
            {
                var filterType = fieldTypeFilter.ToUpper();
                var fieldTypeName = field.Type.ToString().ToUpper();
                if (!fieldTypeName.Contains(filterType) && filterType != "ALL")
                {
                    continue;
                }
            }
            
            var fieldInfo = new FieldInfo
            {
                Index = fieldIndex++,
                Type = field.Type.ToString(),
                Code = field.GetFieldCode(),
                Result = includeResult ? (field.Result ?? "") : null,
                IsLocked = field.IsLocked,
                IsDirty = field.IsDirty
            };
            
            // Get specific field type information
            if (field is FieldHyperlink hyperlinkField)
            {
                fieldInfo.ExtraInfo = $"Address: {hyperlinkField.Address ?? ""}, ScreenTip: {hyperlinkField.ScreenTip ?? ""}";
            }
            else if (field is FieldRef refField)
            {
                fieldInfo.ExtraInfo = $"Bookmark: {refField.BookmarkName ?? ""}";
            }
            else if (field is FieldDate dateField)
            {
                fieldInfo.ExtraInfo = $"Date: {dateField.Result ?? ""}";
            }
            else if (field is FieldPage pageField)
            {
                fieldInfo.ExtraInfo = $"Page Number: {pageField.Result ?? ""}";
            }
            else if (field is FieldNumPages numPagesField)
            {
                fieldInfo.ExtraInfo = $"Total Pages: {numPagesField.Result ?? ""}";
            }
            
            fields.Add(fieldInfo);
        }
        
        var result = new System.Text.StringBuilder();
        result.AppendLine($"文檔中共有 {fields.Count} 個功能變數\n");
        
        if (fields.Count == 0)
        {
            result.AppendLine("未找到功能變數");
            return await Task.FromResult(result.ToString());
        }
        
        result.AppendLine("【功能變數列表】");
        result.AppendLine(new string('-', 80));
        
        foreach (var fieldInfo in fields)
        {
            result.AppendLine($"索引: {fieldInfo.Index}");
            result.AppendLine($"類型: {fieldInfo.Type}");
            
            if (includeCode)
            {
                result.AppendLine($"代碼: {fieldInfo.Code}");
            }
            
            if (includeResult && fieldInfo.Result != null)
            {
                result.AppendLine($"結果: {fieldInfo.Result}");
            }
            
            if (fieldInfo.ExtraInfo != null)
            {
                result.AppendLine($"額外資訊: {fieldInfo.ExtraInfo}");
            }
            
            result.AppendLine($"鎖定: {(fieldInfo.IsLocked ? "是" : "否")}");
            result.AppendLine($"需要更新: {(fieldInfo.IsDirty ? "是" : "否")}");
            result.AppendLine(new string('-', 80));
        }
        
        // Summary by type
        var typeGroups = fields.GroupBy(f => f.Type).OrderBy(g => g.Key);
        result.AppendLine("\n【按類型統計】");
        foreach (var group in typeGroups)
        {
            result.AppendLine($"{group.Key}: {group.Count()} 個");
        }
        
        return await Task.FromResult(result.ToString());
    }
    
    private class FieldInfo
    {
        public int Index { get; set; }
        public string Type { get; set; } = "";
        public string Code { get; set; } = "";
        public string? Result { get; set; }
        public bool IsLocked { get; set; }
        public bool IsDirty { get; set; }
        public string? ExtraInfo { get; set; }
    }
}

