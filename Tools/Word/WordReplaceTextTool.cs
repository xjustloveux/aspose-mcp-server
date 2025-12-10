using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeMcpServer.Tools;

public class WordReplaceTextTool : IAsposeTool
{
    public string Description => "Replace text in a Word document";

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
            find = new
            {
                type = "string",
                description = "Text to find"
            },
            replace = new
            {
                type = "string",
                description = "Replacement text"
            },
            useRegex = new
            {
                type = "boolean",
                description = "Use regex matching (optional)"
            }
        },
        required = new[] { "path", "find", "replace" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var find = arguments?["find"]?.GetValue<string>() ?? throw new ArgumentException("find is required");
        var replace = arguments?["replace"]?.GetValue<string>() ?? throw new ArgumentException("replace is required");
        var useRegex = arguments?["useRegex"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        var options = new FindReplaceOptions();
        if (useRegex)
        {
            doc.Range.Replace(new Regex(find), replace, options);
        }
        else
        {
            doc.Range.Replace(find, replace, options);
        }

        doc.Save(path);

        return await Task.FromResult($"Text replaced in document: {path}");
    }
}

