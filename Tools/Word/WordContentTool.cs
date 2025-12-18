using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

public class WordContentTool : IAsposeTool
{
    public string Description =>
        @"Get content and information from Word documents. Supports 4 operations: get_content, get_content_detailed, get_statistics, get_document_info.

Usage examples:
- Get content: word_content(operation='get_content', path='doc.docx')
- Get detailed content: word_content(operation='get_content_detailed', path='doc.docx', includeHeaders=true, includeFooters=true)
- Get statistics: word_content(operation='get_statistics', path='doc.docx')
- Get document info: word_content(operation='get_document_info', path='doc.docx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get_content': Get document content (required params: path)
- 'get_content_detailed': Get detailed content with structure (required params: path)
- 'get_statistics': Get document statistics (required params: path)
- 'get_document_info': Get document information (required params: path)",
                @enum = new[] { "get_content", "get_content_detailed", "get_statistics", "get_document_info" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            includeHeaders = new
            {
                type = "boolean",
                description = "Include headers (for get_content_detailed, default: true)"
            },
            includeFooters = new
            {
                type = "boolean",
                description = "Include footers (for get_content_detailed, default: true)"
            },
            includeStyles = new
            {
                type = "boolean",
                description = "Include style information (for get_content_detailed, default: true)"
            },
            includeTables = new
            {
                type = "boolean",
                description = "Include table structure details (for get_content_detailed, default: true)"
            },
            includeImages = new
            {
                type = "boolean",
                description = "Include image information (for get_content_detailed, default: true)"
            },
            includeFootnotes = new
            {
                type = "boolean",
                description = "Include footnotes and endnotes in count (for get_statistics, default: true)"
            },
            includeTextboxes = new
            {
                type = "boolean",
                description = "Include text boxes in count (for get_statistics, default: true)"
            },
            includeTabStops = new
            {
                type = "boolean",
                description = "Include tab stops information (for get_document_info, default: true)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "get_content" => await GetContent(arguments),
            "get_content_detailed" => await GetContentDetailed(arguments),
            "get_statistics" => await GetStatistics(arguments),
            "get_document_info" => await GetDocumentInfo(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets document content as plain text
    /// </summary>
    /// <param name="arguments">JSON arguments containing path</param>
    /// <returns>Document content as string</returns>
    private async Task<string> GetContent(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var doc = new Document(path);
        doc.UpdateFields();

        // Get text content - this will show field results (like hyperlink display text) instead of field codes
        var text = doc.Range.Text;
        return await Task.FromResult(text);
    }

    /// <summary>
    ///     Gets detailed document content with structure information
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional includeFormatting</param>
    /// <returns>Formatted string with detailed content</returns>
    private async Task<string> GetContentDetailed(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var includeHeaders = ArgumentHelper.GetBool(arguments, "includeHeaders", true);
        var includeFooters = ArgumentHelper.GetBool(arguments, "includeFooters");
        var includeStyles = ArgumentHelper.GetBool(arguments, "includeStyles", true);
        var includeTables = ArgumentHelper.GetBool(arguments, "includeTables");
        var includeImages = ArgumentHelper.GetBool(arguments, "includeImages");

        var doc = new Document(path);
        var result = new StringBuilder();

        result.AppendLine("=== Document Basic Info ===");
        result.AppendLine($"Pages: {doc.PageCount}");
        result.AppendLine($"Paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}");
        result.AppendLine($"Tables: {doc.GetChildNodes(NodeType.Table, true).Count}");
        result.AppendLine($"Images: {doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Count(s => s.HasImage)}");
        result.AppendLine();

        if (includeHeaders)
        {
            result.AppendLine("=== Headers ===");
            foreach (var section in doc.Sections.Cast<Section>())
            {
                var headerFooter = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (headerFooter != null && !string.IsNullOrWhiteSpace(headerFooter.ToString(SaveFormat.Text)))
                {
                    result.AppendLine($"Section {section.ParentNode.IndexOf(section) + 1} Header:");
                    result.AppendLine(headerFooter.ToString(SaveFormat.Text).Trim());
                    result.AppendLine();
                }
            }
        }

        if (includeFooters)
        {
            result.AppendLine("=== Footers ===");
            foreach (var section in doc.Sections.Cast<Section>())
            {
                var headerFooter = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (headerFooter != null && !string.IsNullOrWhiteSpace(headerFooter.ToString(SaveFormat.Text)))
                {
                    result.AppendLine($"Section {section.ParentNode.IndexOf(section) + 1} Footer:");
                    result.AppendLine(headerFooter.ToString(SaveFormat.Text).Trim());
                    result.AppendLine();
                }
            }
        }

        if (includeStyles)
        {
            result.AppendLine("=== Styles ===");
            var usedStyles = new HashSet<string>();
            foreach (var para in doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>())
                if (para.ParagraphFormat.Style != null)
                    usedStyles.Add(para.ParagraphFormat.Style.Name);
            result.AppendLine($"Used Styles: {string.Join(", ", usedStyles)}");
            result.AppendLine();
        }

        if (includeTables)
        {
            result.AppendLine("=== Tables ===");
            var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            for (var i = 0; i < tables.Count; i++)
                result.AppendLine($"Table {i}: {tables[i].Rows.Count} rows x {tables[i].Rows[0].Cells.Count} columns");
            result.AppendLine();
        }

        if (includeImages)
        {
            result.AppendLine("=== Images ===");
            var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
            result.AppendLine($"Total Images: {shapes.Count}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    ///     Gets document statistics
    /// </summary>
    /// <param name="arguments">JSON arguments containing path</param>
    /// <returns>Formatted string with document statistics</returns>
    private async Task<string> GetStatistics(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        _ = ArgumentHelper.GetBool(arguments, "includeFootnotes", true);
        _ = ArgumentHelper.GetBool(arguments, "includeTextboxes", true);

        var doc = new Document(path);
        doc.UpdateWordCount(true);
        doc.UpdatePageLayout();

        var result = new StringBuilder();
        result.AppendLine("=== Document Statistics ===\n");

        result.AppendLine("【Basic Statistics】");
        result.AppendLine($"Pages: {doc.PageCount}");
        result.AppendLine($"Words: {doc.BuiltInDocumentProperties.Words}");
        result.AppendLine($"Characters (with spaces): {doc.BuiltInDocumentProperties.Characters}");
        result.AppendLine($"Characters (without spaces): {doc.BuiltInDocumentProperties.CharactersWithSpaces}");
        result.AppendLine($"Paragraphs: {doc.BuiltInDocumentProperties.Paragraphs}");
        result.AppendLine($"Lines: {doc.BuiltInDocumentProperties.Lines}");
        result.AppendLine();

        result.AppendLine("【Document Structure】");
        result.AppendLine($"Sections: {doc.Sections.Count}");
        result.AppendLine($"Actual Paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}");
        result.AppendLine($"Tables: {doc.GetChildNodes(NodeType.Table, true).Count}");

        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        var imageCount = shapes.Cast<Shape>().Count(s => s.HasImage);
        var textboxCount = shapes.Cast<Shape>().Count(s => s.ShapeType == ShapeType.TextBox);

        result.AppendLine($"Images: {imageCount}");
        result.AppendLine($"Textboxes: {textboxCount}");
        result.AppendLine();

        result.AppendLine("【Content Elements】");
        result.AppendLine($"Hyperlinks: {doc.Range.Fields.Count(f => f.Type == FieldType.FieldHyperlink)}");
        result.AppendLine($"Bookmarks: {doc.Range.Bookmarks.Count}");
        result.AppendLine($"Comments: {doc.GetChildNodes(NodeType.Comment, true).Count}");
        result.AppendLine($"Fields: {doc.Range.Fields.Count}");
        result.AppendLine();

        if (File.Exists(path))
        {
            var fileInfo = new FileInfo(path);
            result.AppendLine("【File Information】");
            result.AppendLine($"File Size: {FormatFileSize(fileInfo.Length)}");
            result.AppendLine($"Last Modified: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
        }

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    ///     Gets document information
    /// </summary>
    /// <param name="arguments">JSON arguments containing path</param>
    /// <returns>Formatted string with document information</returns>
    private async Task<string> GetDocumentInfo(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        _ = ArgumentHelper.GetBool(arguments, "includeTabStops");

        var doc = new Document(path);
        var result = new StringBuilder();

        result.AppendLine("=== Word Document Detailed Information ===\n");

        result.AppendLine("【File Information】");
        result.AppendLine($"File Format: {doc.OriginalFileName}");
        result.AppendLine($"Sections: {doc.Sections.Count}");
        result.AppendLine($"Paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}");
        result.AppendLine($"Tables: {doc.GetChildNodes(NodeType.Table, true).Count}\n");

        var section = doc.FirstSection;
        if (section != null)
        {
            var pageSetup = section.PageSetup;

            result.AppendLine("【Page Setup】(First Section)");
            result.AppendLine($"Page Width: {pageSetup.PageWidth:F2} pt ({pageSetup.PageWidth / 28.35:F2} cm)");
            result.AppendLine($"Page Height: {pageSetup.PageHeight:F2} pt ({pageSetup.PageHeight / 28.35:F2} cm)");
            result.AppendLine($"Orientation: {pageSetup.Orientation}");
            result.AppendLine();

            result.AppendLine("【Margins】");
            result.AppendLine($"Top: {pageSetup.TopMargin:F2} pt ({pageSetup.TopMargin / 28.35:F2} cm)");
            result.AppendLine($"Bottom: {pageSetup.BottomMargin:F2} pt ({pageSetup.BottomMargin / 28.35:F2} cm)");
            result.AppendLine($"Left: {pageSetup.LeftMargin:F2} pt ({pageSetup.LeftMargin / 28.35:F2} cm)");
            result.AppendLine($"Right: {pageSetup.RightMargin:F2} pt ({pageSetup.RightMargin / 28.35:F2} cm)");
            result.AppendLine();

            result.AppendLine("【Header/Footer Distance】");
            result.AppendLine(
                $"Header Distance: {pageSetup.HeaderDistance:F2} pt ({pageSetup.HeaderDistance / 28.35:F2} cm)");
            result.AppendLine(
                $"Footer Distance: {pageSetup.FooterDistance:F2} pt ({pageSetup.FooterDistance / 28.35:F2} cm)");
            result.AppendLine();

            var contentWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
            result.AppendLine("【Calculated Information】");
            result.AppendLine($"Content Area Width: {contentWidth:F2} pt ({contentWidth / 28.35:F2} cm)");
            result.AppendLine(
                $"Center Position: {pageSetup.PageWidth / 2:F2} pt ({pageSetup.PageWidth / 2 / 28.35:F2} cm)");
            result.AppendLine(
                $"Right Position: {pageSetup.PageWidth - pageSetup.RightMargin:F2} pt ({(pageSetup.PageWidth - pageSetup.RightMargin) / 28.35:F2} cm)");
        }

        return await Task.FromResult(result.ToString());
    }

    private string FormatFileSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB"];
        double len = bytes;
        var order = 0;

        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }

        return $"{len:0.##} {sizes[order]}";
    }
}