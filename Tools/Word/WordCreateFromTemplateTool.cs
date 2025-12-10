using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordCreateFromTemplateTool : IAsposeTool
{
    public string Description => "Create a Word document from a template by replacing placeholders";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            templatePath = new
            {
                type = "string",
                description = "Template document path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path"
            },
            replacements = new
            {
                type = "object",
                description = "Key-value pairs for placeholder replacements, e.g., {\"{{Title}}\": \"My Document\", \"{{Date}}\": \"2025-12-04\"}"
            },
            placeholderStyle = new
            {
                type = "string",
                description = "Placeholder format: doubleCurly ({{key}}), singleCurly ({key}), square ([key]) (default: doubleCurly)",
                @enum = new[] { "doubleCurly", "singleCurly", "square" }
            }
        },
        required = new[] { "templatePath", "outputPath", "replacements" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var templatePath = arguments?["templatePath"]?.GetValue<string>() ?? throw new ArgumentException("templatePath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var placeholderStyle = arguments?["placeholderStyle"]?.GetValue<string>() ?? "doubleCurly";

        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException($"找不到範本文件: {templatePath}");
        }

        // Parse replacements
        var replacements = new Dictionary<string, string>();
        if (arguments?.ContainsKey("replacements") == true)
        {
            try
            {
                var replacementsObj = arguments["replacements"]?.AsObject();
                if (replacementsObj != null)
                {
                    foreach (var kvp in replacementsObj)
                    {
                        var key = kvp.Key;
                        var value = kvp.Value?.GetValue<string>() ?? "";
                        
                        // Ensure key has the correct placeholder format
                        if (!IsValidPlaceholder(key, placeholderStyle))
                        {
                            key = FormatPlaceholder(key, placeholderStyle);
                        }
                        
                        replacements[key] = value;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"無法解析 replacements 參數: {ex.Message}", ex);
            }
        }

        if (replacements.Count == 0)
        {
            throw new ArgumentException("replacements 不能為空");
        }

        // Load template
        var doc = new Document(templatePath);

        // Replace placeholders
        foreach (var kvp in replacements)
        {
            doc.Range.Replace(kvp.Key, kvp.Value, new Aspose.Words.Replacing.FindReplaceOptions());
        }

        // Save output
        doc.Save(outputPath);

        var result = $"成功從範本創建文檔\n";
        result += $"範本: {Path.GetFileName(templatePath)}\n";
        result += $"替換數量: {replacements.Count}\n";
        result += $"替換內容:\n";
        foreach (var kvp in replacements.Take(5))
        {
            result += $"  {kvp.Key} → {kvp.Value}\n";
        }
        if (replacements.Count > 5)
        {
            result += $"  ... 還有 {replacements.Count - 5} 項\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private bool IsValidPlaceholder(string key, string style)
    {
        return style.ToLower() switch
        {
            "singlecurly" => key.StartsWith("{") && key.EndsWith("}"),
            "square" => key.StartsWith("[") && key.EndsWith("]"),
            _ => key.StartsWith("{{") && key.EndsWith("}}")
        };
    }

    private string FormatPlaceholder(string key, string style)
    {
        // Remove any existing placeholder markers
        key = key.Trim('{', '}', '[', ']');

        return style.ToLower() switch
        {
            "singlecurly" => $"{{{key}}}",
            "square" => $"[{key}]",
            _ => $"{{{{{key}}}}}" // doubleCurly
        };
    }
}

