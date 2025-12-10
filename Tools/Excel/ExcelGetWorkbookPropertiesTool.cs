using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetWorkbookPropertiesTool : IAsposeTool
{
    public string Description => "Get workbook properties (metadata) from Excel file";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var workbook = new Workbook(path);
        var props = workbook.BuiltInDocumentProperties;
        var customProps = workbook.CustomDocumentProperties;

        var sb = new StringBuilder();
        sb.AppendLine("Workbook Properties:");
        sb.AppendLine($"  Title: {props.Title ?? "(none)"}");
        sb.AppendLine($"  Subject: {props.Subject ?? "(none)"}");
        sb.AppendLine($"  Author: {props.Author ?? "(none)"}");
        sb.AppendLine($"  Keywords: {props.Keywords ?? "(none)"}");
        sb.AppendLine($"  Comments: {props.Comments ?? "(none)"}");
        sb.AppendLine($"  Category: {props.Category ?? "(none)"}");
        sb.AppendLine($"  Company: {props.Company ?? "(none)"}");
        sb.AppendLine($"  Manager: {props.Manager ?? "(none)"}");
        sb.AppendLine($"  Created: {props.CreatedTime}");
        sb.AppendLine($"  Modified: {props.LastSavedTime}");
        sb.AppendLine($"  Last Saved By: {props.LastSavedBy ?? "(none)"}");
        sb.AppendLine($"  Revision: {props.RevisionNumber}");

        if (customProps.Count > 0)
        {
            sb.AppendLine("\nCustom Properties:");
            for (int i = 0; i < customProps.Count; i++)
            {
                var prop = customProps[i];
                sb.AppendLine($"  {prop.Name}: {prop.Value}");
            }
        }

        sb.AppendLine($"\nTotal Sheets: {workbook.Worksheets.Count}");
        // Note: Workbook protection check may require different API

        return await Task.FromResult(sb.ToString());
    }
}

