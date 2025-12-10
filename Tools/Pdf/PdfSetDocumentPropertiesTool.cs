using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfSetDocumentPropertiesTool : IAsposeTool
{
    public string Description => "Set document properties (metadata) for PDF file";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            title = new
            {
                type = "string",
                description = "Title (optional)"
            },
            author = new
            {
                type = "string",
                description = "Author (optional)"
            },
            subject = new
            {
                type = "string",
                description = "Subject (optional)"
            },
            keywords = new
            {
                type = "string",
                description = "Keywords (optional)"
            },
            creator = new
            {
                type = "string",
                description = "Creator (optional)"
            },
            producer = new
            {
                type = "string",
                description = "Producer (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var title = arguments?["title"]?.GetValue<string>();
        var author = arguments?["author"]?.GetValue<string>();
        var subject = arguments?["subject"]?.GetValue<string>();
        var keywords = arguments?["keywords"]?.GetValue<string>();
        var creator = arguments?["creator"]?.GetValue<string>();
        var producer = arguments?["producer"]?.GetValue<string>();

        using var document = new Document(path);
        var metadata = document.Metadata;

        if (!string.IsNullOrEmpty(title)) metadata["Title"] = title;
        if (!string.IsNullOrEmpty(author)) metadata["Author"] = author;
        if (!string.IsNullOrEmpty(subject)) metadata["Subject"] = subject;
        if (!string.IsNullOrEmpty(keywords)) metadata["Keywords"] = keywords;
        if (!string.IsNullOrEmpty(creator)) metadata["Creator"] = creator;
        if (!string.IsNullOrEmpty(producer)) metadata["Producer"] = producer;

        document.Save(path);
        return await Task.FromResult($"Document properties updated: {path}");
    }
}

