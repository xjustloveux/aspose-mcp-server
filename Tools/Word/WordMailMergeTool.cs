using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for performing mail merge operations on Word document templates
/// </summary>
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var templatePath = ArgumentHelper.GetString(arguments, "templatePath");
        var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
        var data = ArgumentHelper.GetObject(arguments, "data");

        SecurityHelper.ValidateFilePath(templatePath, "templatePath");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(templatePath);

        var fieldNames = data.Select(kvp => kvp.Key).ToArray();
        var fieldValues = data.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.Save(outputPath);

        return await Task.FromResult($"Mail merge completed: {outputPath}");
    }
}