using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;

namespace AsposeMcpServer.Tools;

public class PdfDeleteFormFieldTool : IAsposeTool
{
    public string Description => "Delete a form field from PDF document";

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
            fieldName = new
            {
                type = "string",
                description = "Field name (PartialName or FullName)"
            }
        },
        required = new[] { "path", "fieldName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");

        using var document = new Document(path);
        if (document.Form == null || document.Form.Count == 0)
        {
            throw new ArgumentException("No form fields found in the document");
        }

        Field? fieldToDelete = null;
        foreach (Field field in document.Form)
        {
            if (field.PartialName == fieldName || field.FullName == fieldName)
            {
                fieldToDelete = field;
                break;
            }
        }

        if (fieldToDelete == null)
        {
            throw new ArgumentException($"Field '{fieldName}' not found");
        }

        document.Form.Delete(fieldToDelete);
        document.Save(path);

        return await Task.FromResult($"Form field '{fieldName}' deleted: {path}");
    }
}

