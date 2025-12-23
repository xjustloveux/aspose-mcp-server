using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for getting Word document content, statistics, and document info
///     Supports: get_content, get_content_detailed, get_statistics, get_document_info
/// </summary>
public class WordContentTool : IAsposeTool
{
    public string Description =>
        @"Get Word document content, statistics, and document information. Supports 4 operations: get_content, get_content_detailed, get_statistics, get_document_info.

Usage examples:
- Get content: word_content(operation='get_content', path='doc.docx')
- Get detailed content: word_content(operation='get_content_detailed', path='doc.docx', includeHeaders=true, includeFooters=true)
- Get statistics: word_content(operation='get_statistics', path='doc.docx', includeFootnotes=true)
- Get document info: word_content(operation='get_document_info', path='doc.docx', includeTabStops=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get_content': Get document content as text (required params: path)
- 'get_content_detailed': Get detailed content including headers/footers (required params: path)
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
                description = "Include headers in content (optional, for get_content_detailed, defaults to false)"
            },
            includeFooters = new
            {
                type = "boolean",
                description = "Include footers in content (optional, for get_content_detailed, defaults to false)"
            },
            includeFootnotes = new
            {
                type = "boolean",
                description = "Include footnotes in statistics (optional, for get_statistics, defaults to true)"
            },
            includeTabStops = new
            {
                type = "boolean",
                description = "Include tab stops in document info (optional, for get_document_info, defaults to false)"
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
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);

        return operation.ToLower() switch
        {
            "get_content" => await GetContentAsync(arguments, path),
            "get_content_detailed" => await GetContentDetailedAsync(arguments, path),
            "get_statistics" => await GetStatisticsAsync(arguments, path),
            "get_document_info" => await GetDocumentInfoAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets document content as text
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Document content as string</returns>
    private Task<string> GetContentAsync(JsonObject? _, string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var sb = new StringBuilder();
            sb.AppendLine("=== Document Content ===");
            sb.AppendLine(doc.GetText());
            return sb.ToString();
        });
    }

    /// <summary>
    ///     Gets detailed document content including headers and footers
    /// </summary>
    /// <param name="arguments">JSON arguments containing includeHeaders, includeFooters flags</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Detailed document content as string</returns>
    private Task<string> GetContentDetailedAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var includeHeaders = ArgumentHelper.GetBool(arguments, "includeHeaders", false);
            var includeFooters = ArgumentHelper.GetBool(arguments, "includeFooters", false);

            var doc = new Document(path);
            var sb = new StringBuilder();
            sb.AppendLine("=== Detailed Document Content ===");

            if (includeHeaders)
            {
                sb.AppendLine("\n--- Headers ---");
                foreach (Section section in doc.Sections)
                foreach (HeaderFooter header in section.HeadersFooters)
                    if (header.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                        header.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                        header.HeaderFooterType == HeaderFooterType.HeaderEven)
                    {
                        var headerText = header.GetText();
                        if (!string.IsNullOrWhiteSpace(headerText))
                        {
                            sb.AppendLine($"Section {doc.Sections.IndexOf(section)} - {header.HeaderFooterType}:");
                            sb.AppendLine(headerText);
                        }
                    }
            }

            sb.AppendLine("\n--- Body Content ---");
            sb.AppendLine(doc.GetText());

            if (includeFooters)
            {
                sb.AppendLine("\n--- Footers ---");
                foreach (Section section in doc.Sections)
                foreach (HeaderFooter footer in section.HeadersFooters)
                    if (footer.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                        footer.HeaderFooterType == HeaderFooterType.FooterFirst ||
                        footer.HeaderFooterType == HeaderFooterType.FooterEven)
                    {
                        var footerText = footer.GetText();
                        if (!string.IsNullOrWhiteSpace(footerText))
                        {
                            sb.AppendLine($"Section {doc.Sections.IndexOf(section)} - {footer.HeaderFooterType}:");
                            sb.AppendLine(footerText);
                        }
                    }
            }

            return sb.ToString();
        });
    }

    /// <summary>
    ///     Gets document statistics
    /// </summary>
    /// <param name="arguments">JSON arguments containing includeFootnotes flag</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with document statistics</returns>
    private Task<string> GetStatisticsAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var includeFootnotes = ArgumentHelper.GetBool(arguments, "includeFootnotes", true);

            var doc = new Document(path);
            var sb = new StringBuilder();
            sb.AppendLine("=== Document Statistics ===");

            var stats = doc.BuiltInDocumentProperties;
            sb.AppendLine($"Pages: {stats.Pages}");
            sb.AppendLine($"Words: {stats.Words}");
            sb.AppendLine($"Characters: {stats.Characters}");
            sb.AppendLine($"Characters (with spaces): {stats.CharactersWithSpaces}");
            sb.AppendLine($"Paragraphs: {stats.Paragraphs}");
            sb.AppendLine($"Lines: {stats.Lines}");

            if (includeFootnotes)
            {
                var footnotes = doc.GetChildNodes(NodeType.Footnote, true);
                sb.AppendLine($"Footnotes: {footnotes.Count}");
            }
            else
            {
                sb.AppendLine("Footnotes: excluded from statistics");
            }

            var tables = doc.GetChildNodes(NodeType.Table, true);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            var images = shapes.Count(s => s.HasImage);

            sb.AppendLine($"Tables: {tables.Count}");
            sb.AppendLine($"Images: {images}");
            sb.AppendLine($"Shapes: {shapes.Count}");

            return sb.ToString();
        });
    }

    /// <summary>
    ///     Gets document information
    /// </summary>
    /// <param name="arguments">JSON arguments containing includeTabStops flag</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with document information</returns>
    private Task<string> GetDocumentInfoAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var includeTabStops = ArgumentHelper.GetBool(arguments, "includeTabStops", false);

            var doc = new Document(path);
            var sb = new StringBuilder();
            sb.AppendLine("=== Document Information ===");

            var props = doc.BuiltInDocumentProperties;
            sb.AppendLine($"Title: {props.Title ?? "(none)"}");
            sb.AppendLine($"Author: {props.Author ?? "(none)"}");
            sb.AppendLine($"Subject: {props.Subject ?? "(none)"}");
            sb.AppendLine($"Created: {props.CreatedTime}");
            sb.AppendLine($"Modified: {props.LastSavedTime}");
            sb.AppendLine($"Pages: {props.Pages}");
            sb.AppendLine($"Sections: {doc.Sections.Count}");

            if (includeTabStops)
            {
                sb.AppendLine("\nTab Stops:");
                foreach (Section section in doc.Sections)
                foreach (var para in section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>())
                    if (para.ParagraphFormat.TabStops.Count > 0)
                    {
                        sb.AppendLine(
                            $"  Paragraph {section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList().IndexOf(para)}:");
                        for (var i = 0; i < para.ParagraphFormat.TabStops.Count; i++)
                        {
                            var tabStop = para.ParagraphFormat.TabStops[i];
                            sb.AppendLine($"    Position: {tabStop.Position}, Alignment: {tabStop.Alignment}");
                        }
                    }
            }

            return sb.ToString();
        });
    }
}