using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetWorkbookPropertiesTool : IAsposeTool
{
    public string Description => "Set workbook properties (metadata) for Excel file";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            title = new
            {
                type = "string",
                description = "Title (optional)"
            },
            subject = new
            {
                type = "string",
                description = "Subject (optional)"
            },
            author = new
            {
                type = "string",
                description = "Author (optional)"
            },
            keywords = new
            {
                type = "string",
                description = "Keywords (optional)"
            },
            comments = new
            {
                type = "string",
                description = "Comments (optional)"
            },
            category = new
            {
                type = "string",
                description = "Category (optional)"
            },
            company = new
            {
                type = "string",
                description = "Company (optional)"
            },
            manager = new
            {
                type = "string",
                description = "Manager (optional)"
            },
            customProperties = new
            {
                type = "object",
                description = "Custom properties as key-value pairs (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var title = arguments?["title"]?.GetValue<string>();
        var subject = arguments?["subject"]?.GetValue<string>();
        var author = arguments?["author"]?.GetValue<string>();
        var keywords = arguments?["keywords"]?.GetValue<string>();
        var comments = arguments?["comments"]?.GetValue<string>();
        var category = arguments?["category"]?.GetValue<string>();
        var company = arguments?["company"]?.GetValue<string>();
        var manager = arguments?["manager"]?.GetValue<string>();
        var customProps = arguments?["customProperties"]?.AsObject();

        using var workbook = new Workbook(path);
        var props = workbook.BuiltInDocumentProperties;

        if (!string.IsNullOrEmpty(title)) props.Title = title;
        if (!string.IsNullOrEmpty(subject)) props.Subject = subject;
        if (!string.IsNullOrEmpty(author)) props.Author = author;
        if (!string.IsNullOrEmpty(keywords)) props.Keywords = keywords;
        if (!string.IsNullOrEmpty(comments)) props.Comments = comments;
        if (!string.IsNullOrEmpty(category)) props.Category = category;
        if (!string.IsNullOrEmpty(company)) props.Company = company;
        if (!string.IsNullOrEmpty(manager)) props.Manager = manager;

        if (customProps != null)
        {
            foreach (var kvp in customProps)
            {
                workbook.CustomDocumentProperties.Add(kvp.Key, kvp.Value?.GetValue<string>() ?? "");
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Workbook properties updated: {path}");
    }
}

