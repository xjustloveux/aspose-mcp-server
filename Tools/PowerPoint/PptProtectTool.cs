using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptProtectTool : IAsposeTool
{
    public string Description => "Protect a PowerPoint presentation with password";

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
            password = new
            {
                type = "string",
                description = "Protection password"
            },
            readOnlyRecommended = new
            {
                type = "boolean",
                description = "Mark as read-only recommended (optional, default: false)"
            }
        },
        required = new[] { "path", "password" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var password = arguments?["password"]?.GetValue<string>() ?? throw new ArgumentException("password is required");
        var readOnlyRecommended = arguments?["readOnlyRecommended"]?.GetValue<bool?>() ?? false;

        using var presentation = new Presentation(path);
        presentation.ProtectionManager.Encrypt(password);
        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Presentation protected with password: {path}");
    }
}

