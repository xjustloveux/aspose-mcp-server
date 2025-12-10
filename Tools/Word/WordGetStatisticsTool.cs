using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordGetStatisticsTool : IAsposeTool
{
    public string Description => "Get detailed statistics about a Word document (word count, character count, page count, etc.)";

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
            includeFootnotes = new
            {
                type = "boolean",
                description = "Include footnotes and endnotes in count (default: true)"
            },
            includeTextboxes = new
            {
                type = "boolean",
                description = "Include text boxes in count (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var includeFootnotes = arguments?["includeFootnotes"]?.GetValue<bool>() ?? true;
        var includeTextboxes = arguments?["includeTextboxes"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        
        // Update word count
        doc.UpdateWordCount(true);
        doc.UpdatePageLayout();

        var result = new StringBuilder();
        result.AppendLine("=== 文檔統計資訊 ===\n");

        // Basic counts from built-in properties
        result.AppendLine("【基本統計】");
        result.AppendLine($"頁數: {doc.PageCount}");
        result.AppendLine($"字數: {doc.BuiltInDocumentProperties.Words}");
        result.AppendLine($"字元數（含空格）: {doc.BuiltInDocumentProperties.Characters}");
        result.AppendLine($"字元數（不含空格）: {doc.BuiltInDocumentProperties.CharactersWithSpaces}");
        result.AppendLine($"段落數: {doc.BuiltInDocumentProperties.Paragraphs}");
        result.AppendLine($"行數: {doc.BuiltInDocumentProperties.Lines}");
        result.AppendLine();

        // Document structure counts
        result.AppendLine("【文檔結構】");
        result.AppendLine($"節數: {doc.Sections.Count}");
        result.AppendLine($"實際段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}");
        result.AppendLine($"表格數: {doc.GetChildNodes(NodeType.Table, true).Count}");
        
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        var imageCount = shapes.Cast<Shape>().Count(s => s.HasImage);
        var textboxCount = shapes.Cast<Shape>().Count(s => s.ShapeType == Aspose.Words.Drawing.ShapeType.TextBox);
        
        result.AppendLine($"圖片數: {imageCount}");
        result.AppendLine($"文字框數: {textboxCount}");
        result.AppendLine();

        // Content elements
        result.AppendLine("【內容元素】");
        result.AppendLine($"超連結數: {doc.Range.Fields.Count(f => f.Type == Aspose.Words.Fields.FieldType.FieldHyperlink)}");
        result.AppendLine($"書籤數: {doc.Range.Bookmarks.Count}");
        result.AppendLine($"註釋數: {doc.GetChildNodes(NodeType.Comment, true).Count}");
        result.AppendLine($"欄位數: {doc.Range.Fields.Count}");
        result.AppendLine();

        // Headers and footers
        var headerCount = 0;
        var footerCount = 0;
        foreach (Section section in doc.Sections)
        {
            if (section.HeadersFooters[HeaderFooterType.HeaderPrimary] != null &&
                !string.IsNullOrWhiteSpace(section.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText()))
                headerCount++;
            
            if (section.HeadersFooters[HeaderFooterType.FooterPrimary] != null &&
                !string.IsNullOrWhiteSpace(section.HeadersFooters[HeaderFooterType.FooterPrimary].GetText()))
                footerCount++;
        }
        
        result.AppendLine("【頁首頁尾】");
        result.AppendLine($"含頁首的節數: {headerCount}");
        result.AppendLine($"含頁尾的節數: {footerCount}");
        result.AppendLine();

        // List information
        var listParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Where(p => p.ListFormat != null && p.ListFormat.IsListItem)
            .ToList();
        
        result.AppendLine("【列表】");
        result.AppendLine($"列表項數: {listParagraphs.Count}");
        result.AppendLine();

        // Style usage
        var usedStyles = new HashSet<string>();
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.Style != null)
                usedStyles.Add(para.ParagraphFormat.Style.Name);
        }
        
        result.AppendLine("【樣式】");
        result.AppendLine($"文檔中定義的樣式數: {doc.Styles.Count}");
        result.AppendLine($"實際使用的樣式數: {usedStyles.Count}");
        result.AppendLine();

        // File information
        if (System.IO.File.Exists(path))
        {
            var fileInfo = new System.IO.FileInfo(path);
            result.AppendLine("【檔案資訊】");
            result.AppendLine($"檔案大小: {FormatFileSize(fileInfo.Length)}");
            result.AppendLine($"最後修改時間: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
            result.AppendLine();
        }

        // Document properties
        if (doc.BuiltInDocumentProperties.LastSavedTime.Year > 1900)
        {
            result.AppendLine("【文檔屬性】");
            if (!string.IsNullOrEmpty(doc.BuiltInDocumentProperties.Author))
                result.AppendLine($"作者: {doc.BuiltInDocumentProperties.Author}");
            if (!string.IsNullOrEmpty(doc.BuiltInDocumentProperties.Title))
                result.AppendLine($"標題: {doc.BuiltInDocumentProperties.Title}");
            if (!string.IsNullOrEmpty(doc.BuiltInDocumentProperties.Subject))
                result.AppendLine($"主旨: {doc.BuiltInDocumentProperties.Subject}");
            if (doc.BuiltInDocumentProperties.CreatedTime.Year > 1900)
                result.AppendLine($"創建時間: {doc.BuiltInDocumentProperties.CreatedTime:yyyy-MM-dd HH:mm:ss}");
            if (doc.BuiltInDocumentProperties.LastSavedTime.Year > 1900)
                result.AppendLine($"上次儲存: {doc.BuiltInDocumentProperties.LastSavedTime:yyyy-MM-dd HH:mm:ss}");
        }

        return await Task.FromResult(result.ToString());
    }

    private string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        
        return $"{len:0.##} {sizes[order]}";
    }
}

