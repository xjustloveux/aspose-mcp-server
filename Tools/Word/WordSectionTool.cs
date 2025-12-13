using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Word sections (insert, delete, get info)
/// Merges: WordInsertSectionTool, WordDeleteSectionTool, WordGetSectionsTool, WordGetSectionsInfoTool
/// </summary>
public class WordSectionTool : IAsposeTool
{
    public string Description => @"Manage Word document sections. Supports 3 operations: insert, delete, get.

Usage examples:
- Insert section: word_section(operation='insert', path='doc.docx', sectionBreakType='NextPage', insertAtParagraphIndex=5)
- Delete section: word_section(operation='delete', path='doc.docx', sectionIndex=1)
- Get sections: word_section(operation='get', path='doc.docx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'insert': Insert a section break (required params: path, sectionBreakType)
- 'delete': Delete a section (required params: path, sectionIndex)
- 'get': Get all sections info (required params: path)",
                @enum = new[] { "insert", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for insert/delete operations)"
            },
            sectionBreakType = new
            {
                type = "string",
                description = "Section break type: 'NextPage', 'Continuous', 'EvenPage', 'OddPage' (required for insert operation)",
                @enum = new[] { "NextPage", "Continuous", "EvenPage", "OddPage" }
            },
            insertAtParagraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert section break after (0-based, optional, default: end of document, for insert operation)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, required for delete operation, optional for get operation)"
            },
            sectionIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of section indices to delete (0-based, optional, overrides sectionIndex, for delete operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        return operation.ToLower() switch
        {
            "insert" => await InsertSectionAsync(arguments, path),
            "delete" => await DeleteSectionAsync(arguments, path),
            "get" => await GetSectionsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> InsertSectionAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sectionBreakType = arguments?["sectionBreakType"]?.GetValue<string>() ?? throw new ArgumentException("sectionBreakType is required for insert operation");
        var insertAtParagraphIndex = arguments?["insertAtParagraphIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        var breakType = sectionBreakType switch
        {
            "NextPage" => SectionStart.NewPage,
            "Continuous" => SectionStart.Continuous,
            "EvenPage" => SectionStart.EvenPage,
            "OddPage" => SectionStart.OddPage,
            _ => SectionStart.NewPage
        };

        if (insertAtParagraphIndex.HasValue)
        {
            if (insertAtParagraphIndex.Value == -1)
            {
                // insertAtParagraphIndex=-1 means document end
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.SectionBreakContinuous);
                builder.CurrentSection.PageSetup.SectionStart = breakType;
            }
            else
            {
                var actualSectionIndex = sectionIndex ?? 0;
                if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
                {
                    actualSectionIndex = 0;
                }

                var section = doc.Sections[actualSectionIndex];
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                
                if (insertAtParagraphIndex.Value < 0 || insertAtParagraphIndex.Value >= paragraphs.Count)
                {
                    throw new ArgumentException($"insertAtParagraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");
                }

                var para = paragraphs[insertAtParagraphIndex.Value];
                builder.MoveTo(para);
                builder.InsertBreak(BreakType.SectionBreakContinuous);
                builder.CurrentSection.PageSetup.SectionStart = breakType;
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.CurrentSection.PageSetup.SectionStart = breakType;
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Section break inserted ({sectionBreakType}): {outputPath}");
    }

    private async Task<string> DeleteSectionAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var sectionIndicesArray = arguments?["sectionIndices"]?.AsArray();

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        if (doc.Sections.Count <= 1)
        {
            throw new ArgumentException("Cannot delete the last section. Document must have at least one section.");
        }

        List<int> sectionsToDelete;
        if (sectionIndicesArray != null && sectionIndicesArray.Count > 0)
        {
            sectionsToDelete = sectionIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue).Select(s => s!.Value).OrderByDescending(s => s).ToList();
        }
        else if (sectionIndex.HasValue)
        {
            sectionsToDelete = new List<int> { sectionIndex.Value };
        }
        else
        {
            throw new ArgumentException("Either sectionIndex or sectionIndices must be provided for delete operation");
        }

        foreach (var idx in sectionsToDelete)
        {
            if (idx < 0 || idx >= doc.Sections.Count)
            {
                continue;
            }
            if (doc.Sections.Count <= 1)
            {
                break; // Don't delete the last section
            }
            doc.Sections.RemoveAt(idx);
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Deleted {sectionsToDelete.Count} section(s). Remaining sections: {doc.Sections.Count}. Output: {outputPath}");
    }

    private async Task<string> GetSectionsAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var result = new StringBuilder();

        result.AppendLine("=== 文檔節資訊 ===\n");
        result.AppendLine($"總節數: {doc.Sections.Count}\n");

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"節索引 {sectionIndex.Value} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
            }
            
            var section = doc.Sections[sectionIndex.Value];
            AppendSectionInfo(result, section, sectionIndex.Value);
        }
        else
        {
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                var section = doc.Sections[i];
                AppendSectionInfo(result, section, i);
                if (i < doc.Sections.Count - 1)
                {
                    result.AppendLine();
                }
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSectionInfo(StringBuilder result, Section section, int index)
    {
        result.AppendLine($"【節 {index}】");
        
        var pageSetup = section.PageSetup;
        result.AppendLine($"頁面設置:");
        result.AppendLine($"  紙張大小: {pageSetup.PaperSize}");
        result.AppendLine($"  方向: {pageSetup.Orientation}");
        result.AppendLine($"  上邊距: {pageSetup.TopMargin} 點");
        result.AppendLine($"  下邊距: {pageSetup.BottomMargin} 點");
        result.AppendLine($"  左邊距: {pageSetup.LeftMargin} 點");
        result.AppendLine($"  右邊距: {pageSetup.RightMargin} 點");
        result.AppendLine($"  頁眉距離: {pageSetup.HeaderDistance} 點");
        result.AppendLine($"  頁尾距離: {pageSetup.FooterDistance} 點");
        result.AppendLine($"  頁碼起始: {(pageSetup.RestartPageNumbering ? pageSetup.PageStartingNumber.ToString() : "繼承上一節")}");
        result.AppendLine($"  不同首頁: {pageSetup.DifferentFirstPageHeaderFooter}");
        result.AppendLine($"  不同奇偶頁: {pageSetup.OddAndEvenPagesHeaderFooter}");
        result.AppendLine($"  分欄數: {pageSetup.TextColumns.Count}");
        
        result.AppendLine();
        result.AppendLine($"內容統計:");
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true);
        var tables = section.Body.GetChildNodes(NodeType.Table, true);
        var shapes = section.Body.GetChildNodes(NodeType.Shape, true);
        result.AppendLine($"  段落數: {paragraphs.Count}");
        result.AppendLine($"  表格數: {tables.Count}");
        result.AppendLine($"  形狀數: {shapes.Count}");
        
        result.AppendLine();
        result.AppendLine($"頁眉頁尾:");
        var headerCount = 0;
        var footerCount = 0;
        foreach (HeaderFooter hf in section.HeadersFooters)
        {
            if (hf.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                hf.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                hf.HeaderFooterType == HeaderFooterType.HeaderEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    headerCount++;
            }
            else if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                     hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                     hf.HeaderFooterType == HeaderFooterType.FooterEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    footerCount++;
            }
        }
        result.AppendLine($"  頁眉數: {headerCount}");
        result.AppendLine($"  頁尾數: {footerCount}");
    }
}

