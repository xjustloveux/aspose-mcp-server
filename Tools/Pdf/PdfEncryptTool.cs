using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfEncryptTool : IAsposeTool
{
    public string Description => "Encrypt a PDF document with password";

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
            userPassword = new
            {
                type = "string",
                description = "User password (for opening the document)"
            },
            ownerPassword = new
            {
                type = "string",
                description = "Owner password (for permissions)"
            }
        },
        required = new[] { "path", "userPassword", "ownerPassword" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var userPassword = arguments?["userPassword"]?.GetValue<string>() ?? throw new ArgumentException("userPassword is required");
        var ownerPassword = arguments?["ownerPassword"]?.GetValue<string>() ?? throw new ArgumentException("ownerPassword is required");

        using var document = new Document(path);
        document.Encrypt(userPassword, ownerPassword, Permissions.PrintDocument | Permissions.ModifyContent, CryptoAlgorithm.AESx256);
        document.Save(path);

        return await Task.FromResult($"PDF encrypted with password: {path}");
    }
}

