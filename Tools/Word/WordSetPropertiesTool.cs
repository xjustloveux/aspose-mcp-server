using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Properties;

namespace AsposeMcpServer.Tools;

public class WordSetPropertiesTool : IAsposeTool
{
    public string Description => "Set document properties (metadata) in a Word document";

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
            title = new
            {
                type = "string",
                description = "Document title"
            },
            subject = new
            {
                type = "string",
                description = "Document subject"
            },
            author = new
            {
                type = "string",
                description = "Document author"
            },
            company = new
            {
                type = "string",
                description = "Company name"
            },
            manager = new
            {
                type = "string",
                description = "Manager name"
            },
            category = new
            {
                type = "string",
                description = "Document category"
            },
            keywords = new
            {
                type = "array",
                description = "Keywords array",
                items = new { type = "string" }
            },
            comments = new
            {
                type = "string",
                description = "Comments"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var title = arguments?["title"]?.GetValue<string>();
        var subject = arguments?["subject"]?.GetValue<string>();
        var author = arguments?["author"]?.GetValue<string>();
        var company = arguments?["company"]?.GetValue<string>();
        var manager = arguments?["manager"]?.GetValue<string>();
        var category = arguments?["category"]?.GetValue<string>();
        var comments = arguments?["comments"]?.GetValue<string>();

        var doc = new Document(path);
        var builtInProps = doc.BuiltInDocumentProperties;

        var changesCount = 0;

        if (!string.IsNullOrEmpty(title))
        {
            builtInProps.Title = title;
            changesCount++;
        }

        if (!string.IsNullOrEmpty(subject))
        {
            builtInProps.Subject = subject;
            changesCount++;
        }

        if (!string.IsNullOrEmpty(author))
        {
            builtInProps.Author = author;
            changesCount++;
        }

        if (!string.IsNullOrEmpty(company))
        {
            builtInProps.Company = company;
            changesCount++;
        }

        if (!string.IsNullOrEmpty(manager))
        {
            builtInProps.Manager = manager;
            changesCount++;
        }

        if (!string.IsNullOrEmpty(category))
        {
            builtInProps.Category = category;
            changesCount++;
        }

        if (!string.IsNullOrEmpty(comments))
        {
            builtInProps.Comments = comments;
            changesCount++;
        }

        // Parse keywords
        if (arguments?.ContainsKey("keywords") == true)
        {
            try
            {
                var keywordsArray = arguments["keywords"]?.AsArray();
                if (keywordsArray != null && keywordsArray.Count > 0)
                {
                    var keywords = new List<string>();
                    foreach (var item in keywordsArray)
                    {
                        var keyword = item?.GetValue<string>();
                        if (!string.IsNullOrEmpty(keyword))
                        {
                            keywords.Add(keyword);
                        }
                    }
                    if (keywords.Count > 0)
                    {
                        builtInProps.Keywords = string.Join(", ", keywords);
                        changesCount++;
                    }
                }
            }
            catch { }
        }

        if (changesCount == 0)
        {
            return await Task.FromResult("警告: 未提供任何屬性更新");
        }

        doc.Save(outputPath);

        var result = $"成功設定文檔屬性\n";
        result += $"更新項目數: {changesCount}\n";
        if (!string.IsNullOrEmpty(title)) result += $"標題: {title}\n";
        if (!string.IsNullOrEmpty(subject)) result += $"主旨: {subject}\n";
        if (!string.IsNullOrEmpty(author)) result += $"作者: {author}\n";
        if (!string.IsNullOrEmpty(company)) result += $"公司: {company}\n";
        if (!string.IsNullOrEmpty(category)) result += $"類別: {category}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

