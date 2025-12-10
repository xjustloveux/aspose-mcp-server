using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordAddImageTool : IAsposeTool
{
    public string Description => "Add an image to a Word document (supports size, alignment, text wrapping, caption)";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            imagePath = new
            {
                type = "string",
                description = "Image file path"
            },
            width = new
            {
                type = "number",
                description = "Image width in points (optional)"
            },
            height = new
            {
                type = "number",
                description = "Image height in points (optional)"
            },
            alignment = new
            {
                type = "string",
                description = "Horizontal alignment: left, center, right (default: left)",
                @enum = new[] { "left", "center", "right" }
            },
            textWrapping = new
            {
                type = "string",
                description = "Text wrapping: inline, square, tight, through, topAndBottom, none (default: inline)",
                @enum = new[] { "inline", "square", "tight", "through", "topAndBottom", "none" }
            },
            caption = new
            {
                type = "string",
                description = "Image caption text (optional)"
            },
            captionPosition = new
            {
                type = "string",
                description = "Caption position: above, below (default: below)",
                @enum = new[] { "above", "below" }
            }
        },
        required = new[] { "path", "imagePath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();
        var alignment = arguments?["alignment"]?.GetValue<string>() ?? "left";
        var textWrapping = arguments?["textWrapping"]?.GetValue<string>() ?? "inline";
        var caption = arguments?["caption"]?.GetValue<string>();
        var captionPosition = arguments?["captionPosition"]?.GetValue<string>() ?? "below";

        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"找不到圖片文件: {imagePath}");
        }

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        // Add caption above if specified
        if (!string.IsNullOrEmpty(caption) && captionPosition == "above")
        {
            builder.ParagraphFormat.Alignment = GetAlignment(alignment);
            builder.Font.Italic = true;
            builder.Writeln(caption);
            builder.Font.Italic = false;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }

        // Set paragraph alignment for image
        builder.ParagraphFormat.Alignment = GetAlignment(alignment);

        // Insert image
        Shape shape;
        if (textWrapping == "inline")
        {
            // Inline image
            shape = builder.InsertImage(imagePath);
        }
        else
        {
            // Floating image
            shape = builder.InsertImage(imagePath);
            shape.WrapType = GetWrapType(textWrapping);
            
            // Set relative horizontal position for center/right alignment
            if (alignment == "center")
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.HorizontalAlignment = HorizontalAlignment.Center;
            }
            else if (alignment == "right")
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.HorizontalAlignment = HorizontalAlignment.Right;
            }
        }

        // Set size if specified
        if (width.HasValue)
            shape.Width = width.Value;
        
        if (height.HasValue)
            shape.Height = height.Value;

        // Reset paragraph alignment
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Add caption below if specified
        if (!string.IsNullOrEmpty(caption) && captionPosition == "below")
        {
            builder.ParagraphFormat.Alignment = GetAlignment(alignment);
            builder.Font.Italic = true;
            builder.Writeln(caption);
            builder.Font.Italic = false;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }

        doc.Save(outputPath);

        var result = $"成功添加圖片\n";
        result += $"圖片: {Path.GetFileName(imagePath)}\n";
        if (width.HasValue || height.HasValue)
        {
            result += $"尺寸: {(width.HasValue ? width.Value.ToString() : "auto")} x {(height.HasValue ? height.Value.ToString() : "auto")} pt\n";
        }
        result += $"對齊: {alignment}\n";
        result += $"文繞圖: {textWrapping}\n";
        if (!string.IsNullOrEmpty(caption))
        {
            result += $"圖片說明: {caption} ({captionPosition})\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }

    private WrapType GetWrapType(string wrapType)
    {
        return wrapType.ToLower() switch
        {
            "square" => WrapType.Square,
            "tight" => WrapType.Tight,
            "through" => WrapType.Through,
            "topandbottom" => WrapType.TopBottom,
            "none" => WrapType.None,
            _ => WrapType.Inline
        };
    }
}

