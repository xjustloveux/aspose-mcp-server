using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Cells;
using Aspose.Slides;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class ConvertDocumentTool : IAsposeTool
{
    public string Description => "Convert documents between various formats (auto-detect source type)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPath = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (extension determines output format)"
            }
        },
        required = new[] { "inputPath", "outputPath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateFilePath(inputPath, "inputPath");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var inputExtension = Path.GetExtension(inputPath).ToLower();
        var outputExtension = Path.GetExtension(outputPath).ToLower();

        // Detect document type and convert
        if (IsWordDocument(inputExtension))
        {
            var doc = new Document(inputPath);
            var saveFormat = GetWordSaveFormat(outputExtension);
            doc.Save(outputPath, saveFormat);
        }
        else if (IsExcelDocument(inputExtension))
        {
            using var workbook = new Workbook(inputPath);
            var saveFormat = GetExcelSaveFormat(outputExtension);
            workbook.Save(outputPath, saveFormat);
        }
        else if (IsPresentationDocument(inputExtension))
        {
            using var presentation = new Presentation(inputPath);
            var saveFormat = GetPresentationSaveFormat(outputExtension);
            presentation.Save(outputPath, saveFormat);
        }
        else
        {
            throw new ArgumentException($"Unsupported input format: {inputExtension}");
        }

        return await Task.FromResult($"Document converted from {inputPath} to {outputPath}");
    }

    private bool IsWordDocument(string extension)
    {
        return extension is ".doc" or ".docx" or ".rtf" or ".odt" or ".txt";
    }

    private bool IsExcelDocument(string extension)
    {
        return extension is ".xls" or ".xlsx" or ".csv" or ".ods";
    }

    private bool IsPresentationDocument(string extension)
    {
        return extension is ".ppt" or ".pptx" or ".odp";
    }

    private Aspose.Words.SaveFormat GetWordSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Words.SaveFormat.Pdf,
            ".docx" => Aspose.Words.SaveFormat.Docx,
            ".doc" => Aspose.Words.SaveFormat.Doc,
            ".rtf" => Aspose.Words.SaveFormat.Rtf,
            ".html" => Aspose.Words.SaveFormat.Html,
            ".txt" => Aspose.Words.SaveFormat.Text,
            ".odt" => Aspose.Words.SaveFormat.Odt,
            _ => throw new ArgumentException($"Unsupported output format: {extension}")
        };
    }

    private Aspose.Cells.SaveFormat GetExcelSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Cells.SaveFormat.Pdf,
            ".xlsx" => Aspose.Cells.SaveFormat.Xlsx,
            ".xls" => Aspose.Cells.SaveFormat.Excel97To2003,
            ".csv" => Aspose.Cells.SaveFormat.Csv,
            ".html" => Aspose.Cells.SaveFormat.Html,
            ".ods" => Aspose.Cells.SaveFormat.Ods,
            _ => throw new ArgumentException($"Unsupported output format: {extension}")
        };
    }

    private Aspose.Slides.Export.SaveFormat GetPresentationSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Slides.Export.SaveFormat.Pdf,
            ".pptx" => Aspose.Slides.Export.SaveFormat.Pptx,
            ".ppt" => Aspose.Slides.Export.SaveFormat.Ppt,
            ".html" => Aspose.Slides.Export.SaveFormat.Html,
            ".odp" => Aspose.Slides.Export.SaveFormat.Odp,
            _ => throw new ArgumentException($"Unsupported output format: {extension}")
        };
    }
}

