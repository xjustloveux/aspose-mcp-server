using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetPageSizeTool : IAsposeTool
{
    public string Description => "Set page size for section(s) in Word document";

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
            width = new
            {
                type = "number",
                description = "Page width in points"
            },
            height = new
            {
                type = "number",
                description = "Page height in points"
            },
            paperSize = new
            {
                type = "string",
                description = "Predefined paper size: 'A4', 'Letter', 'Legal', 'A3', 'A5' (optional, overrides width/height)",
                @enum = new[] { "A4", "Letter", "Legal", "A3", "A5" }
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided applies to all sections)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();
        var paperSize = arguments?["paperSize"]?.GetValue<string>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        List<int> sectionsToUpdate;

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }
            sectionsToUpdate = new List<int> { sectionIndex.Value };
        }
        else
        {
            sectionsToUpdate = Enumerable.Range(0, doc.Sections.Count).ToList();
        }

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;
            
            if (!string.IsNullOrEmpty(paperSize))
            {
                pageSetup.PaperSize = paperSize.ToUpper() switch
                {
                    "A4" => PaperSize.A4,
                    "LETTER" => PaperSize.Letter,
                    "LEGAL" => PaperSize.Legal,
                    "A3" => PaperSize.A3,
                    "A5" => PaperSize.A5,
                    _ => PaperSize.A4
                };
            }
            else if (width.HasValue && height.HasValue)
            {
                pageSetup.PageWidth = width.Value;
                pageSetup.PageHeight = height.Value;
            }
            else
            {
                throw new ArgumentException("Either paperSize or both width and height must be provided");
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Page size updated for {sectionsToUpdate.Count} section(s): {path}");
    }
}

