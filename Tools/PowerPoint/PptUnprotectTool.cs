using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptUnprotectTool : IAsposeTool
{
    public string Description => "Remove password protection from a PowerPoint presentation";

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
                description = "Current protection password"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, if not provided overwrites input)"
            }
        },
        required = new[] { "path", "password" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var password = arguments?["password"]?.GetValue<string>() ?? throw new ArgumentException("password is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        
        if (presentation.ProtectionManager.IsWriteProtected)
        {
            if (!presentation.ProtectionManager.CheckWriteProtection(password))
            {
                throw new ArgumentException("Invalid password");
            }
            presentation.ProtectionManager.RemoveWriteProtection();
        }

        var savePath = outputPath ?? path;
        presentation.Save(savePath, SaveFormat.Pptx);

        return await Task.FromResult($"Password protection removed: {savePath}");
    }
}

