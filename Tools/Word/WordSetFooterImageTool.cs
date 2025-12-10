using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordSetFooterImageTool : IAsposeTool
{
    public string Description => "Set image for footer in a Word document (fine-grained control)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            imagePath = new
            {
                type = "string",
                description = "Path to image file"
            },
            alignment = new
            {
                type = "string",
                description = "Image alignment: left, center, right (default: left)",
                @enum = new[] { "left", "center", "right" }
            },
            width = new
            {
                type = "number",
                description = "Image width in points (default: 20)"
            },
            height = new
            {
                type = "number",
                description = "Image height in points (optional, maintains aspect ratio if not specified)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to apply to all sections"
            },
            removeExisting = new
            {
                type = "boolean",
                description = "Remove existing images before adding new one (default: true)"
            }
        },
        required = new[] { "path", "imagePath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
        var alignment = arguments?["alignment"]?.GetValue<string>() ?? "left";
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var removeExisting = arguments?["removeExisting"]?.GetValue<bool>() ?? true;

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"找不到圖片文件: {imagePath}");

        var doc = new Document(path);
        
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : new[] { doc.Sections[sectionIndex] };

        foreach (Section section in sections)
        {
            var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer == null)
            {
                footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                section.HeadersFooters.Add(footer);
            }

            // Remove existing images if requested
            if (removeExisting)
            {
                var existingShapes = footer.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                    .Where(s => s.HasImage).ToList();
                foreach (var existingShape in existingShapes)
                {
                    existingShape.Remove();
                }
            }

            // Add image
            var imagePara = new Paragraph(doc);
            imagePara.ParagraphFormat.Alignment = GetAlignment(alignment);
            
            // First append paragraph to footer, then insert image
            footer.AppendChild(imagePara);
            
            var builder = new DocumentBuilder(doc);
            builder.MoveTo(imagePara);
            
            var imageShape = builder.InsertImage(imagePath);
            imageShape.Width = width.HasValue ? width.Value : 20.0;
            if (height.HasValue)
                imageShape.Height = height.Value;
            else
                imageShape.AspectRatioLocked = true;
        }

        doc.Save(outputPath);
        
        var sectionsDesc = sectionIndex == -1 ? "所有節" : $"第 {sectionIndex} 節";
        
        return await Task.FromResult($"成功設定頁尾圖片於 {sectionsDesc}");
    }
    
    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }
}

