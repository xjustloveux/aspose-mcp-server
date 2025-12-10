using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetProtectionTool : IAsposeTool
{
    public string Description => "Get protection information for a presentation";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var presentation = new Presentation(path);
        var protection = presentation.ProtectionManager;
        var sb = new StringBuilder();

        sb.AppendLine("=== Protection Information ===");
        sb.AppendLine($"IsEncrypted: {protection.IsEncrypted}");
        sb.AppendLine($"IsWriteProtected: {protection.IsWriteProtected}");

        return await Task.FromResult(sb.ToString());
    }
}

