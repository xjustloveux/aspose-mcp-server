using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetHeadersFootersTool : IAsposeTool
{
    public string Description => "Get all headers and footers from a Word document";

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
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided returns all sections)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var result = new StringBuilder();

        result.AppendLine("=== 文檔頁眉頁腳資訊 ===\n");
        result.AppendLine($"總節數: {doc.Sections.Count}\n");

        var sections = sectionIndex.HasValue 
            ? new[] { doc.Sections[sectionIndex.Value] }
            : doc.Sections.Cast<Section>().ToArray();

        if (sectionIndex.HasValue && (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
        {
            throw new ArgumentException($"節索引 {sectionIndex.Value} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
        }

        for (int i = 0; i < sections.Length; i++)
        {
            var section = sections[i];
            var actualIndex = sectionIndex ?? i;
            
            result.AppendLine($"【節 {actualIndex}】");
            
            // Headers
            result.AppendLine("頁眉:");
            var headerTypes = new[]
            {
                (HeaderFooterType.HeaderPrimary, "主要頁眉"),
                (HeaderFooterType.HeaderFirst, "首頁頁眉"),
                (HeaderFooterType.HeaderEven, "偶數頁頁眉")
            };
            
            bool hasHeader = false;
            foreach (var (type, name) in headerTypes)
            {
                var header = section.HeadersFooters[type];
                if (header != null)
                {
                    var headerText = header.GetText().Trim();
                    if (!string.IsNullOrEmpty(headerText))
                    {
                        result.AppendLine($"  {name}:");
                        result.AppendLine($"    {headerText.Replace("\n", "\n    ")}");
                        hasHeader = true;
                    }
                }
            }
            
            if (!hasHeader)
            {
                result.AppendLine("  (無頁眉)");
            }
            
            result.AppendLine();
            
            // Footers
            result.AppendLine("頁尾:");
            var footerTypes = new[]
            {
                (HeaderFooterType.FooterPrimary, "主要頁尾"),
                (HeaderFooterType.FooterFirst, "首頁頁尾"),
                (HeaderFooterType.FooterEven, "偶數頁頁尾")
            };
            
            bool hasFooter = false;
            foreach (var (type, name) in footerTypes)
            {
                var footer = section.HeadersFooters[type];
                if (footer != null)
                {
                    var footerText = footer.GetText().Trim();
                    if (!string.IsNullOrEmpty(footerText))
                    {
                        result.AppendLine($"  {name}:");
                        result.AppendLine($"    {footerText.Replace("\n", "\n    ")}");
                        hasFooter = true;
                    }
                }
            }
            
            if (!hasFooter)
            {
                result.AppendLine("  (無頁尾)");
            }
            
            if (i < sections.Length - 1)
            {
                result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }
}

