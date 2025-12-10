using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetDocumentPropertiesTool : IAsposeTool
{
    public string Description => "Set document properties (metadata)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
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

        using var presentation = new Presentation(path);
        var props = presentation.DocumentProperties;
        var changes = new List<string>();

        if (title != null)
        {
            props.Title = title;
            changes.Add("Title");
        }
        if (subject != null)
        {
            props.Subject = subject;
            changes.Add("Subject");
        }
        if (author != null)
        {
            props.Author = author;
            changes.Add("Author");
        }
        if (keywords != null)
        {
            props.Keywords = keywords;
            changes.Add("Keywords");
        }
        if (comments != null)
        {
            props.Comments = comments;
            changes.Add("Comments");
        }
        if (category != null)
        {
            props.Category = category;
            changes.Add("Category");
        }
        if (company != null)
        {
            props.Company = company;
            changes.Add("Company");
        }
        if (manager != null)
        {
            props.Manager = manager;
            changes.Add("Manager");
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Document properties updated: {string.Join(", ", changes)} - {path}");
    }
}

