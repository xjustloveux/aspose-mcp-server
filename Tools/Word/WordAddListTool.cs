using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeMcpServer.Tools;

public class WordAddListTool : IAsposeTool
{
    public string Description => "Add a bulleted or numbered list to a Word document";

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
            items = new
            {
                type = "array",
                description = "List items (strings or objects with 'text' and 'level')",
                items = new
                {
                    oneOf = new object[]
                    {
                        new { type = "string" },
                        new
                        {
                            type = "object",
                            properties = new
                            {
                                text = new { type = "string" },
                                level = new { type = "number", description = "Indent level (0-8, default: 0)" }
                            }
                        }
                    }
                }
            },
            listType = new
            {
                type = "string",
                description = "List type: bullet, number, custom (default: bullet)",
                @enum = new[] { "bullet", "number", "custom" }
            },
            bulletChar = new
            {
                type = "string",
                description = "Custom bullet character (only for custom type, e.g., '●', '■', '▪')"
            },
            numberFormat = new
            {
                type = "string",
                description = "Number format for numbered lists: arabic (1,2,3), roman (I,II,III), letter (A,B,C) (default: arabic)",
                @enum = new[] { "arabic", "roman", "letter" }
            }
        },
        required = new[] { "path", "items" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var listType = arguments?["listType"]?.GetValue<string>() ?? "bullet";
        var bulletChar = arguments?["bulletChar"]?.GetValue<string>() ?? "●";
        var numberFormat = arguments?["numberFormat"]?.GetValue<string>() ?? "arabic";

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        // Parse items
        var items = ParseItems(arguments?["items"]);

        if (items.Count == 0)
        {
            throw new ArgumentException("items 不能為空");
        }

        // Create list
        var list = doc.Lists.Add(listType == "number" ? ListTemplate.NumberDefault : ListTemplate.BulletDefault);

        // Customize list format if needed
        if (listType == "custom" && !string.IsNullOrEmpty(bulletChar))
        {
            list.ListLevels[0].NumberFormat = bulletChar;
            list.ListLevels[0].NumberStyle = NumberStyle.Bullet;
        }
        else if (listType == "number")
        {
            var numStyle = numberFormat.ToLower() switch
            {
                "roman" => NumberStyle.UppercaseRoman,
                "letter" => NumberStyle.UppercaseLetter,
                _ => NumberStyle.Arabic
            };

            for (int i = 0; i < list.ListLevels.Count; i++)
            {
                list.ListLevels[i].NumberStyle = numStyle;
            }
        }

        // Add items
        foreach (var item in items)
        {
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = Math.Min(item.level, 8); // Max 9 levels (0-8)
            builder.Writeln(item.text);
        }

        builder.ListFormat.RemoveNumbers();

        doc.Save(outputPath);

        var result = $"成功添加清單\n";
        result += $"類型: {listType}\n";
        if (listType == "custom") result += $"項目符號: {bulletChar}\n";
        if (listType == "number") result += $"數字格式: {numberFormat}\n";
        result += $"項目數: {items.Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private List<(string text, int level)> ParseItems(JsonNode? itemsNode)
    {
        var items = new List<(string text, int level)>();

        if (itemsNode == null)
            return items;

        try
        {
            var itemsArray = itemsNode.AsArray();
            foreach (var item in itemsArray)
            {
                if (item is JsonValue jsonValue)
                {
                    // Simple string item
                    items.Add((jsonValue.GetValue<string>(), 0));
                }
                else if (item is JsonObject jsonObj)
                {
                    // Object with text and level
                    var text = jsonObj["text"]?.GetValue<string>() ?? "";
                    var level = jsonObj["level"]?.GetValue<int>() ?? 0;
                    items.Add((text, level));
                }
            }
        }
        catch
        {
            // Return empty list on parse error
        }

        return items;
    }
}

