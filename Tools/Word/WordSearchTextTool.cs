using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSearchTextTool : IAsposeTool
{
    public string Description => "Search for text in a Word document and return all matches with location and context";

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
            searchText = new
            {
                type = "string",
                description = "Text to search for"
            },
            useRegex = new
            {
                type = "boolean",
                description = "Use regex matching (default: false)"
            },
            caseSensitive = new
            {
                type = "boolean",
                description = "Case sensitive search (default: false)"
            },
            maxResults = new
            {
                type = "number",
                description = "Maximum number of results to return (default: 50)"
            },
            contextLength = new
            {
                type = "number",
                description = "Number of characters to show before and after match for context (default: 50)"
            }
        },
        required = new[] { "path", "searchText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var searchText = arguments?["searchText"]?.GetValue<string>() ?? throw new ArgumentException("searchText is required");
        var useRegex = arguments?["useRegex"]?.GetValue<bool>() ?? false;
        var caseSensitive = arguments?["caseSensitive"]?.GetValue<bool>() ?? false;
        var maxResults = arguments?["maxResults"]?.GetValue<int>() ?? 50;
        var contextLength = arguments?["contextLength"]?.GetValue<int>() ?? 50;

        var doc = new Document(path);
        var result = new StringBuilder();
        var matches = new List<(string text, int paragraphIndex, string context)>();

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        for (int i = 0; i < paragraphs.Count && matches.Count < maxResults; i++)
        {
            var para = paragraphs[i] as Paragraph;
            if (para == null) continue;
            
            var paraText = para.GetText();
            
            // Search for matches
            if (useRegex)
            {
                var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
                var regex = new Regex(searchText, options);
                var regexMatches = regex.Matches(paraText);
                
                foreach (Match match in regexMatches)
                {
                    if (matches.Count >= maxResults) break;
                    
                    var context = GetContext(paraText, match.Index, match.Length, contextLength);
                    matches.Add((match.Value, i, context));
                }
            }
            else
            {
                var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                int index = 0;
                
                while ((index = paraText.IndexOf(searchText, index, comparison)) != -1)
                {
                    if (matches.Count >= maxResults) break;
                    
                    var context = GetContext(paraText, index, searchText.Length, contextLength);
                    matches.Add((searchText, i, context));
                    index += searchText.Length;
                }
            }
        }

        // Format results
        result.AppendLine($"=== 搜尋結果 ===");
        result.AppendLine($"搜尋文字: {searchText}");
        result.AppendLine($"使用正則表達式: {(useRegex ? "是" : "否")}");
        result.AppendLine($"區分大小寫: {(caseSensitive ? "是" : "否")}");
        result.AppendLine($"找到 {matches.Count} 個匹配項{(matches.Count >= maxResults ? $" (限制前 {maxResults} 個)" : "")}\n");

        if (matches.Count == 0)
        {
            result.AppendLine("未找到匹配的文字");
        }
        else
        {
            for (int i = 0; i < matches.Count; i++)
            {
                var match = matches[i];
                result.AppendLine($"匹配 #{i + 1}:");
                result.AppendLine($"  位置: 段落 #{match.paragraphIndex}");
                result.AppendLine($"  匹配文字: {match.text}");
                result.AppendLine($"  上下文: ...{match.context}...");
                result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private string GetContext(string text, int matchIndex, int matchLength, int contextLength)
    {
        int start = Math.Max(0, matchIndex - contextLength);
        int end = Math.Min(text.Length, matchIndex + matchLength + contextLength);
        
        var context = text.Substring(start, end - start);
        
        // Clean up line breaks for display
        context = context.Replace("\r", "").Replace("\n", " ").Trim();
        
        // Highlight the match
        int highlightStart = matchIndex - start;
        int highlightEnd = highlightStart + matchLength;
        
        if (highlightStart >= 0 && highlightEnd <= context.Length)
        {
            context = context.Substring(0, highlightStart) + 
                     "【" + context.Substring(highlightStart, matchLength) + "】" + 
                     context.Substring(highlightEnd);
        }
        
        return context;
    }
}

