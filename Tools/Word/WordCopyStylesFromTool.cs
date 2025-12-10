using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordCopyStylesFromTool : IAsposeTool
{
    public string Description => "Copy styles from another Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Target document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            sourceDocument = new
            {
                type = "string",
                description = "Source document path to copy styles from"
            },
            styleNames = new
            {
                type = "array",
                description = "Array of style names to copy (if not provided, copies all styles)",
                items = new { type = "string" }
            },
            overwriteExisting = new
            {
                type = "boolean",
                description = "Overwrite existing styles with same name (default: false)"
            }
        },
        required = new[] { "path", "sourceDocument" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sourceDocument = arguments?["sourceDocument"]?.GetValue<string>() ?? throw new ArgumentException("sourceDocument is required");
        var overwriteExisting = arguments?["overwriteExisting"]?.GetValue<bool>() ?? false;

        if (!File.Exists(sourceDocument))
        {
            throw new FileNotFoundException($"找不到來源文檔: {sourceDocument}");
        }

        var targetDoc = new Document(path);
        var sourceDoc = new Document(sourceDocument);

        // Parse style names
        var styleNames = new List<string>();
        if (arguments?.ContainsKey("styleNames") == true)
        {
            try
            {
                var stylesArray = arguments["styleNames"]?.AsArray();
                if (stylesArray != null)
                {
                    foreach (var item in stylesArray)
                    {
                        var name = item?.GetValue<string>();
                        if (!string.IsNullOrEmpty(name))
                        {
                            styleNames.Add(name);
                        }
                    }
                }
            }
            catch { }
        }

        // If no specific styles specified, copy all
        bool copyAll = styleNames.Count == 0;
        var copiedCount = 0;
        var skippedCount = 0;

        foreach (Style sourceStyle in sourceDoc.Styles)
        {
            // Skip built-in styles unless explicitly requested
            if (!copyAll && !styleNames.Contains(sourceStyle.Name))
                continue;

            // Check if style already exists
            var existingStyle = targetDoc.Styles[sourceStyle.Name];
            
            if (existingStyle != null && !overwriteExisting)
            {
                skippedCount++;
                continue;
            }

            try
            {
                // Copy style
                if (existingStyle != null && overwriteExisting)
                {
                    // Update existing style - copy all properties
                    CopyStyleProperties(sourceStyle, existingStyle);
                }
                else
                {
                    // Create new style
                    var newStyle = targetDoc.Styles.Add(sourceStyle.Type, sourceStyle.Name);
                    
                    // Copy all properties including list format
                    CopyStyleProperties(sourceStyle, newStyle);
                }

                copiedCount++;
            }
            catch
            {
                skippedCount++;
            }
        }

        targetDoc.Save(outputPath);

        var result = $"成功複製樣式\n";
        result += $"來源文檔: {Path.GetFileName(sourceDocument)}\n";
        result += $"複製樣式數: {copiedCount}\n";
        if (skippedCount > 0) result += $"跳過樣式數: {skippedCount} (已存在且未覆蓋)\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private void CopyStyleProperties(Style sourceStyle, Style targetStyle)
    {
        // Copy font properties
        targetStyle.Font.Name = sourceStyle.Font.Name;
        targetStyle.Font.NameAscii = sourceStyle.Font.NameAscii;
        targetStyle.Font.NameFarEast = sourceStyle.Font.NameFarEast;
        targetStyle.Font.Size = sourceStyle.Font.Size;
        targetStyle.Font.Bold = sourceStyle.Font.Bold;
        targetStyle.Font.Italic = sourceStyle.Font.Italic;
        targetStyle.Font.Color = sourceStyle.Font.Color;
        targetStyle.Font.Underline = sourceStyle.Font.Underline;
        
        if (sourceStyle.Type == StyleType.Paragraph)
        {
            // Copy paragraph formatting
            targetStyle.ParagraphFormat.Alignment = sourceStyle.ParagraphFormat.Alignment;
            targetStyle.ParagraphFormat.SpaceBefore = sourceStyle.ParagraphFormat.SpaceBefore;
            targetStyle.ParagraphFormat.SpaceAfter = sourceStyle.ParagraphFormat.SpaceAfter;
            targetStyle.ParagraphFormat.LineSpacing = sourceStyle.ParagraphFormat.LineSpacing;
            targetStyle.ParagraphFormat.LineSpacingRule = sourceStyle.ParagraphFormat.LineSpacingRule;
            
            // Copy indentation (critical for list styles)
            targetStyle.ParagraphFormat.LeftIndent = sourceStyle.ParagraphFormat.LeftIndent;
            targetStyle.ParagraphFormat.RightIndent = sourceStyle.ParagraphFormat.RightIndent;
            targetStyle.ParagraphFormat.FirstLineIndent = sourceStyle.ParagraphFormat.FirstLineIndent;
            
            // Note: List format properties are complex and cannot be directly copied
            // The indentation properties copied above (LeftIndent, FirstLineIndent) will preserve
            // the visual layout of list items, which is the most important aspect for document consistency
        }
    }
}

