using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordGetFieldDetailTool : IAsposeTool
{
    public string Description => "Get detailed information about a specific field in a Word document";

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
            fieldIndex = new
            {
                type = "number",
                description = "Field index (0-based)"
            }
        },
        required = new[] { "path", "fieldIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int>() ?? throw new ArgumentException("fieldIndex is required");

        var doc = new Document(path);
        
        var fields = doc.Range.Fields.ToList();
        
        if (fieldIndex < 0 || fieldIndex >= fields.Count)
        {
            throw new ArgumentException($"功能變數索引 {fieldIndex} 超出範圍 (文檔共有 {fields.Count} 個功能變數)");
        }
        
        var field = fields[fieldIndex];
        var result = new System.Text.StringBuilder();
        
        result.AppendLine("【功能變數詳細資訊】");
        result.AppendLine(new string('=', 80));
        
        // Basic information
        result.AppendLine($"索引: {fieldIndex}");
        result.AppendLine($"類型: {field.Type}");
        result.AppendLine($"類型代碼: {(int)field.Type}");
        result.AppendLine($"代碼: {field.GetFieldCode()}");
        result.AppendLine($"結果: {field.Result ?? "(無結果)"}");
        result.AppendLine($"鎖定: {(field.IsLocked ? "是" : "否")}");
        result.AppendLine($"需要更新: {(field.IsDirty ? "是" : "否")}");
        
        // Position information - FieldStart and FieldEnd don't expose Start property directly
        if (field.Start != null)
        {
            result.AppendLine("起始節點: FieldStart");
        }
        if (field.End != null)
        {
            result.AppendLine("結束節點: FieldEnd");
        }
        
        // Field-specific information
        result.AppendLine("\n【欄位特定資訊】");
        
        if (field is FieldHyperlink hyperlinkField)
        {
            result.AppendLine("類型: 超連結");
            result.AppendLine($"地址: {hyperlinkField.Address ?? "(無)"}");
            result.AppendLine($"顯示文字: {hyperlinkField.Result ?? "(無)"}");
            result.AppendLine($"提示文字: {hyperlinkField.ScreenTip ?? "(無)"}");
            result.AppendLine($"目標框架: {hyperlinkField.Target ?? "(無)"}");
        }
        else if (field is FieldRef refField)
        {
            result.AppendLine("類型: 交叉引用");
            result.AppendLine($"書籤名稱: {refField.BookmarkName ?? "(無)"}");
            result.AppendLine($"包含編號: {refField.IncludeNoteOrComment}");
            result.AppendLine($"插入相對位置: {refField.InsertRelativePosition}");
        }
        else if (field is FieldDate dateField)
        {
            result.AppendLine("類型: 日期");
            result.AppendLine($"日期格式: {dateField.GetFieldCode()}");
            result.AppendLine($"日期值: {dateField.Result ?? "(無)"}");
        }
        else if (field is FieldTime timeField)
        {
            result.AppendLine("類型: 時間");
            result.AppendLine($"時間格式: {timeField.GetFieldCode()}");
            result.AppendLine($"時間值: {timeField.Result ?? "(無)"}");
        }
        else if (field is FieldPage pageField)
        {
            result.AppendLine("類型: 頁碼");
            result.AppendLine($"頁碼格式: {pageField.GetFieldCode()}");
            result.AppendLine($"頁碼值: {pageField.Result ?? "(無)"}");
        }
        else if (field is FieldNumPages numPagesField)
        {
            result.AppendLine("類型: 總頁數");
            result.AppendLine($"總頁數值: {numPagesField.Result ?? "(無)"}");
        }
        else if (field is FieldDocProperty docPropField)
        {
            result.AppendLine("類型: 文檔屬性");
            result.AppendLine($"欄位代碼: {docPropField.GetFieldCode()}");
            result.AppendLine($"屬性值: {docPropField.Result ?? "(無)"}");
        }
        else if (field is FieldMergeField mergeField)
        {
            result.AppendLine("類型: 郵件合併欄位");
            result.AppendLine($"欄位名稱: {mergeField.FieldName ?? "(無)"}");
        }
        else if (field is FieldIf ifField)
        {
            result.AppendLine("類型: 條件欄位");
            result.AppendLine($"條件表達式: {ifField.GetFieldCode()}");
        }
        else if (field is FieldSeq seqField)
        {
            result.AppendLine("類型: 序號欄位");
            result.AppendLine($"標識符: {seqField.SequenceIdentifier ?? "(無)"}");
            result.AppendLine($"序號值: {seqField.Result ?? "(無)"}");
        }
        else if (field is FieldToc tocField)
        {
            result.AppendLine("類型: 目錄");
            result.AppendLine($"目錄代碼: {tocField.GetFieldCode()}");
        }
        else
        {
            result.AppendLine($"類型: {field.Type}");
            result.AppendLine($"欄位代碼: {field.GetFieldCode()}");
        }
        
        // Parent node information
        if (field.Start != null && field.Start.ParentNode != null)
        {
            result.AppendLine("\n【父節點資訊】");
            result.AppendLine($"節點類型: {field.Start.ParentNode.NodeType}");
            
            if (field.Start.ParentNode is Paragraph para)
            {
                result.AppendLine($"段落索引: {doc.GetChildNodes(NodeType.Paragraph, true).IndexOf(para)}");
            }
        }
        
        result.AppendLine(new string('=', 80));
        
        return await Task.FromResult(result.ToString());
    }
}

