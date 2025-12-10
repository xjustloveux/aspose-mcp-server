using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordProtectTool : IAsposeTool
{
    public string Description => "Protect a Word document with password";

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
            password = new
            {
                type = "string",
                description = "Protection password"
            },
            protectionType = new
            {
                type = "string",
                description = "Protection type: ReadOnly, AllowOnlyComments, AllowOnlyFormFields, AllowOnlyRevisions"
            }
        },
        required = new[] { "path", "password" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var password = arguments?["password"]?.GetValue<string>() ?? throw new ArgumentException("password is required");
        var protectionTypeStr = arguments?["protectionType"]?.GetValue<string>() ?? "ReadOnly";

        var doc = new Document(path);

        var protectionType = protectionTypeStr switch
        {
            "ReadOnly" => ProtectionType.ReadOnly,
            "AllowOnlyComments" => ProtectionType.AllowOnlyComments,
            "AllowOnlyFormFields" => ProtectionType.AllowOnlyFormFields,
            "AllowOnlyRevisions" => ProtectionType.AllowOnlyRevisions,
            _ => ProtectionType.ReadOnly
        };

        doc.Protect(protectionType, password);
        doc.Save(path);

        return await Task.FromResult($"Document protected with {protectionType}: {path}");
    }
}

