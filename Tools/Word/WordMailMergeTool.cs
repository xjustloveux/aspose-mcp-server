using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.MailMerging;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordMailMergeTool : IAsposeTool
{
    public string Description => @"Perform mail merge on a Word document template.

Usage examples:
- Mail merge: word_mail_merge(templatePath='template.docx', outputPath='output.docx', data={'name':'John','address':'123 Main St'})";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            templatePath = new
            {
                type = "string",
                description = "Template file path (required)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (required)"
            },
            data = new
            {
                type = "object",
                description = "Key-value pairs for mail merge fields"
            }
        },
        required = new[] { "templatePath", "outputPath", "data" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var templatePath = arguments?["templatePath"]?.GetValue<string>() ?? throw new ArgumentException("templatePath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var data = arguments?["data"]?.AsObject() ?? throw new ArgumentException("data is required");

        SecurityHelper.ValidateFilePath(templatePath, "templatePath");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(templatePath);

        var fieldNames = data.Select(kvp => kvp.Key).ToArray();
        var fieldValues = data.Select(kvp => kvp.Value?.ToString() ?? "").ToArray();

        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.Save(outputPath);

        return await Task.FromResult($"Mail merge completed: {outputPath}");
    }
}

