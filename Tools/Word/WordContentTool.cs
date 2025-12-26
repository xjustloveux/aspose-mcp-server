using System.Text;
using System.Text.Json;
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
            },
            maxChars = new
            {
                type = "number",
                description =
                    "Maximum characters to return (optional, for get_content/get_content_detailed). Use for large documents to avoid token overflow. Default: no limit."
            },
            offset = new
            {
                type = "number",
                description =
                    "Character offset to start reading from (optional, for get_content/get_content_detailed, default: 0). Use with maxChars for paginated reading."
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
            "get_content" => await GetContentAsync(path, arguments),
            "get_content_detailed" => await GetContentDetailedAsync(path, arguments),
            "get_statistics" => await GetStatisticsAsync(path, arguments),
            "get_document_info" => await GetDocumentInfoAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets document content as text
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing optional maxChars, offset</param>
    /// <returns>Document content as string</returns>
    private Task<string> GetContentAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var maxChars = ArgumentHelper.GetIntNullable(arguments, "maxChars");
            var offset = ArgumentHelper.GetInt(arguments, "offset", 0);

            var doc = new Document(path);
            var fullText = CleanText(doc.GetText());
            var totalLength = fullText.Length;

            string content;
            var hasMore = false;
            if (offset >= totalLength)
            {
                content = "";
            }
            else if (maxChars.HasValue)
            {
                var endIndex = Math.Min(offset + maxChars.Value, totalLength);
                content = fullText.Substring(offset, endIndex - offset);
                hasMore = endIndex < totalLength;
            }
            else
            {
                content = offset > 0 ? fullText.Substring(offset) : fullText;
            }

            var sb = new StringBuilder();
            sb.AppendLine("=== Document Content ===");
            if (maxChars.HasValue || offset > 0)
            {
                sb.AppendLine($"[Showing chars {offset} to {offset + content.Length} of {totalLength}]");
                if (hasMore)
                    sb.AppendLine($"[More content available, use offset={offset + content.Length} to continue]");
            }

            sb.AppendLine(content);
            return sb.ToString();
        });
    }

    /// <summary>
    ///     Gets detailed document content including headers and footers
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing includeHeaders, includeFooters flags</param>
    /// <returns>Detailed document content as string</returns>
    private Task<string> GetContentDetailedAsync(string path, JsonObject? arguments)
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
                foreach (var section in doc.Sections.Cast<Section>())
                foreach (var header in section.HeadersFooters.Cast<HeaderFooter>())
                    if (header.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                        header.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                        header.HeaderFooterType == HeaderFooterType.HeaderEven)
                    {
                        var headerText = CleanText(header.GetText());
                        if (!string.IsNullOrWhiteSpace(headerText))
                        {
                            sb.AppendLine($"Section {doc.Sections.IndexOf(section)} - {header.HeaderFooterType}:");
                            sb.AppendLine(headerText);
                        }
                    }
            }

            sb.AppendLine("\n--- Body Content ---");
            foreach (var section in doc.Sections.Cast<Section>())
            {
                var bodyText = CleanText(section.Body.GetText());
                if (!string.IsNullOrWhiteSpace(bodyText))
                    sb.AppendLine(bodyText);
            }

            if (includeFooters)
            {
                sb.AppendLine("\n--- Footers ---");
                foreach (var section in doc.Sections.Cast<Section>())
                foreach (var footer in section.HeadersFooters.Cast<HeaderFooter>())
                    if (footer.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                        footer.HeaderFooterType == HeaderFooterType.FooterFirst ||
                        footer.HeaderFooterType == HeaderFooterType.FooterEven)
                    {
                        var footerText = CleanText(footer.GetText());
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
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing includeFootnotes flag</param>
    /// <returns>JSON formatted string with document statistics for better LLM processing</returns>
    private Task<string> GetStatisticsAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var includeFootnotes = ArgumentHelper.GetBool(arguments, "includeFootnotes", true);

            var doc = new Document(path);

            doc.UpdateWordCount();

            var stats = doc.BuiltInDocumentProperties;

            var tables = doc.GetChildNodes(NodeType.Table, true);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            var images = shapes.Count(s => s.HasImage);

            // Build JSON response for better LLM processing
            var result = new
            {
                pages = stats.Pages,
                words = stats.Words,
                characters = stats.Characters,
                charactersWithSpaces = stats.CharactersWithSpaces,
                paragraphs = stats.Paragraphs,
                lines = stats.Lines,
                footnotes = includeFootnotes ? doc.GetChildNodes(NodeType.Footnote, true).Count : (int?)null,
                footnotesIncluded = includeFootnotes,
                tables = tables.Count,
                images,
                shapes = shapes.Count,
                statisticsUpdated = true
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Gets document information
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing includeTabStops flag</param>
    /// <returns>JSON formatted string with document information for better LLM processing</returns>
    private Task<string> GetDocumentInfoAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var includeTabStops = ArgumentHelper.GetBool(arguments, "includeTabStops", false);

            var doc = new Document(path);
            var props = doc.BuiltInDocumentProperties;

            // Build tab stops list if requested
            List<object>? tabStopsList = null;
            if (includeTabStops)
            {
                tabStopsList = new List<object>();
                var sectionIndex = 0;
                foreach (var section in doc.Sections.Cast<Section>())
                {
                    var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    for (var paraIndex = 0; paraIndex < paragraphs.Count; paraIndex++)
                    {
                        var para = paragraphs[paraIndex];
                        if (para.ParagraphFormat.TabStops.Count > 0)
                        {
                            var stops = new List<object>();
                            for (var i = 0; i < para.ParagraphFormat.TabStops.Count; i++)
                            {
                                var tabStop = para.ParagraphFormat.TabStops[i];
                                stops.Add(new
                                {
                                    position = tabStop.Position,
                                    alignment = tabStop.Alignment.ToString()
                                });
                            }

                            tabStopsList.Add(new
                            {
                                sectionIndex,
                                paragraphIndex = paraIndex,
                                tabStops = stops
                            });
                        }
                    }

                    sectionIndex++;
                }
            }

            // Build JSON response for better LLM processing
            var result = new
            {
                title = props.Title,
                author = props.Author,
                subject = props.Subject,
                created = props.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss"),
                modified = props.LastSavedTime.ToString("yyyy-MM-dd HH:mm:ss"),
                pages = props.Pages,
                sections = doc.Sections.Count,
                tabStopsIncluded = includeTabStops,
                tabStops = tabStopsList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Cleans text by removing control characters and normalizing whitespace
    /// </summary>
    /// <param name="text">Raw text from document</param>
    /// <returns>Cleaned text suitable for LLM processing</returns>
    private static string CleanText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var sb = new StringBuilder();
        var lastWasNewline = false;
        var lastWasSpace = false;

        foreach (var c in text)
        {
            // Skip control characters except newline and tab
            if (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t')
                continue;

            // Convert \r\n or \r to \n
            if (c == '\r')
                continue;

            // Handle newlines - collapse multiple newlines into max 2
            if (c == '\n')
            {
                if (!lastWasNewline)
                {
                    sb.Append('\n');
                    lastWasNewline = true;
                }
                else
                {
                    // Allow one blank line (2 newlines max)
                    // Already have one newline, add one more for blank line
                    if (sb is { Length: >= 1 } && sb[^1] == '\n' && (sb is not { Length: >= 2 } || sb[^2] != '\n'))
                        sb.Append('\n');
                }

                lastWasSpace = false;
                continue;
            }

            // Handle spaces - collapse multiple spaces into one
            if (c == ' ' || c == '\t')
            {
                if (!lastWasSpace && !lastWasNewline)
                {
                    sb.Append(' ');
                    lastWasSpace = true;
                }

                continue;
            }

            // Regular character
            sb.Append(c);
            lastWasNewline = false;
            lastWasSpace = false;
        }

        return sb.ToString().Trim();
    }
}