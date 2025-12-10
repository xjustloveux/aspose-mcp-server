using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordEditImageTool : IAsposeTool
{
    public string Description => "Edit properties of an existing image in a Word document";

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
            imageIndex = new
            {
                type = "number",
                description = "Image index (0-based, from word_get_content_detailed or word_extract_images)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to search all sections"
            },
            // Size properties
            width = new
            {
                type = "number",
                description = "Image width in points"
            },
            height = new
            {
                type = "number",
                description = "Image height in points"
            },
            aspectRatioLocked = new
            {
                type = "boolean",
                description = "Lock aspect ratio (default: true)"
            },
            // Alignment properties
            alignment = new
            {
                type = "string",
                description = "Image alignment: left, center, right (for inline images)",
                @enum = new[] { "left", "center", "right" }
            },
            // Wrapping properties
            textWrapping = new
            {
                type = "string",
                description = "Text wrapping: inline, square, tight, through, none",
                @enum = new[] { "inline", "square", "tight", "through", "none" }
            },
            // Position properties (for floating images)
            horizontalAlignment = new
            {
                type = "string",
                description = "Horizontal alignment: left, center, right (for floating images)",
                @enum = new[] { "left", "center", "right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Vertical alignment: top, center, bottom (for floating images)",
                @enum = new[] { "top", "center", "bottom" }
            },
            // Alternative text
            alternativeText = new
            {
                type = "string",
                description = "Alternative text for accessibility"
            },
            title = new
            {
                type = "string",
                description = "Image title"
            }
        },
        required = new[] { "path", "imageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        var doc = new Document(path);
        
        // Get all images (shapes with images) from document
        List<Shape> allImages = new List<Shape>();
        
        if (sectionIndex == -1)
        {
            // Search all sections
            foreach (Section section in doc.Sections)
            {
                var shapes = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
                allImages.AddRange(shapes);
            }
        }
        else
        {
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
            
            var section = doc.Sections[sectionIndex];
            allImages = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        }
        
        if (imageIndex >= allImages.Count)
            throw new ArgumentException($"Image index {imageIndex} out of range (total images: {allImages.Count})");
        
        var shape = allImages[imageIndex];
        
        // Apply size properties
        if (arguments?["width"] != null)
        {
            var width = arguments["width"]?.GetValue<double>();
            if (width.HasValue)
                shape.Width = width.Value;
        }
        
        if (arguments?["height"] != null)
        {
            var height = arguments["height"]?.GetValue<double>();
            if (height.HasValue)
                shape.Height = height.Value;
        }
        
        if (arguments?["aspectRatioLocked"] != null)
        {
            shape.AspectRatioLocked = arguments["aspectRatioLocked"]?.GetValue<bool>() ?? true;
        }
        
        // Apply alignment (for inline images)
        if (arguments?["alignment"] != null)
        {
            var alignment = arguments["alignment"]?.GetValue<string>() ?? "left";
            var parentPara = shape.ParentNode as Paragraph;
            if (parentPara != null)
            {
                parentPara.ParagraphFormat.Alignment = GetAlignment(alignment);
            }
        }
        
        // Apply text wrapping
        if (arguments?["textWrapping"] != null)
        {
            var textWrapping = arguments["textWrapping"]?.GetValue<string>() ?? "inline";
            shape.WrapType = GetWrapType(textWrapping);
            
            // If switching to floating, set position
            if (textWrapping != "inline")
            {
                if (arguments?["horizontalAlignment"] != null)
                {
                    var hAlign = arguments["horizontalAlignment"]?.GetValue<string>() ?? "left";
                    shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                    shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);
                }
                
                if (arguments?["verticalAlignment"] != null)
                {
                    var vAlign = arguments["verticalAlignment"]?.GetValue<string>() ?? "top";
                    shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                    shape.VerticalAlignment = GetVerticalAlignment(vAlign);
                }
            }
        }
        else
        {
            // Apply position properties for floating images
            if (shape.WrapType != WrapType.Inline)
            {
                if (arguments?["horizontalAlignment"] != null)
                {
                    var hAlign = arguments["horizontalAlignment"]?.GetValue<string>() ?? "left";
                    shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                    shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);
                }
                
                if (arguments?["verticalAlignment"] != null)
                {
                    var vAlign = arguments["verticalAlignment"]?.GetValue<string>() ?? "top";
                    shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                    shape.VerticalAlignment = GetVerticalAlignment(vAlign);
                }
            }
        }
        
        // Apply alternative text
        if (arguments?["alternativeText"] != null)
        {
            var altText = arguments["alternativeText"]?.GetValue<string>();
            shape.AlternativeText = altText ?? "";
        }
        
        // Apply title
        if (arguments?["title"] != null)
        {
            var title = arguments["title"]?.GetValue<string>();
            shape.Title = title ?? "";
        }
        
        doc.Save(outputPath);
        
        var changes = new List<string>();
        if (arguments?["width"] != null) changes.Add($"寬度：{arguments["width"]?.GetValue<double>()}");
        if (arguments?["height"] != null) changes.Add($"高度：{arguments["height"]?.GetValue<double>()}");
        if (arguments?["alignment"] != null) changes.Add($"對齊：{arguments["alignment"]?.GetValue<string>()}");
        if (arguments?["textWrapping"] != null) changes.Add($"文字環繞：{arguments["textWrapping"]?.GetValue<string>()}");
        
        var changesDesc = changes.Count > 0 ? string.Join("、", changes) : "屬性";
        
        return await Task.FromResult($"成功編輯圖片 {imageIndex} 的{changesDesc}");
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
    
    private WrapType GetWrapType(string wrapping)
    {
        return wrapping.ToLower() switch
        {
            "inline" => WrapType.Inline,
            "square" => WrapType.Square,
            "tight" => WrapType.Tight,
            "through" => WrapType.Through,
            "none" => WrapType.None,
            _ => WrapType.Inline
        };
    }
    
    private HorizontalAlignment GetHorizontalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => HorizontalAlignment.Left,
            "center" => HorizontalAlignment.Center,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Left
        };
    }
    
    private VerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => VerticalAlignment.Top,
            "center" => VerticalAlignment.Center,
            "bottom" => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Top
        };
    }
}

