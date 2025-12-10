using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfGetBookmarksTool : IAsposeTool
{
    public string Description => "Get all bookmarks from PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var document = new Document(path);
        var sb = new StringBuilder();

        sb.AppendLine($"Bookmarks ({document.Outlines.Count}):");
        sb.AppendLine();

        for (int i = 0; i < document.Outlines.Count; i++)
        {
            var bookmark = document.Outlines[i];
            sb.AppendLine($"[{i}] {bookmark.Title}");
            sb.AppendLine($"    Level: {bookmark.Level}");
            if (bookmark.Action != null && bookmark.Action is Aspose.Pdf.Annotations.GoToAction goToAction)
            {
                sb.AppendLine($"    Destination: (available)");
            }
            else if (bookmark.Destination != null)
            {
                sb.AppendLine($"    Destination: (available)");
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

