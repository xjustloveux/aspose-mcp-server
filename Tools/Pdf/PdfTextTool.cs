using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfTextTool : IAsposeTool
{
    public string Description => "Manage text in PDF documents (add, edit, extract)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add, edit, extract",
                @enum = new[] { "add", "edit", "extract" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input for add/edit, required for extract)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add, edit, extract)"
            },
            text = new
            {
                type = "string",
                description = "Text to add (required for add)"
            },
            x = new
            {
                type = "number",
                description = "X position (for add, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (for add, default: 700)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (for add, default: 'Arial')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (for add, default: 12)"
            },
            oldText = new
            {
                type = "string",
                description = "Text to replace (required for edit)"
            },
            newText = new
            {
                type = "string",
                description = "New text (required for edit)"
            },
            replaceAll = new
            {
                type = "boolean",
                description = "Replace all occurrences (for edit, default: false)"
            },
            includeFontInfo = new
            {
                type = "boolean",
                description = "Include font information (for extract, default: false)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add" => await AddText(arguments),
            "edit" => await EditText(arguments),
            "extract" => await ExtractText(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddText(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var x = arguments?["x"]?.GetValue<double>() ?? 100;
        var y = arguments?["y"]?.GetValue<double>() ?? 700;
        var fontName = arguments?["fontName"]?.GetValue<string>() ?? "Arial";
        var fontSize = arguments?["fontSize"]?.GetValue<double>() ?? 12;

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var textFragment = new TextFragment(text);
        textFragment.Position = new Position(x, y);
        textFragment.TextState.FontSize = (float)fontSize;
        textFragment.TextState.Font = FontRepository.FindFont(fontName);

        var textBuilder = new TextBuilder(page);
        textBuilder.AppendText(textFragment);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully added text to page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> EditText(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var oldText = arguments?["oldText"]?.GetValue<string>() ?? throw new ArgumentException("oldText is required");
        var newText = arguments?["newText"]?.GetValue<string>() ?? throw new ArgumentException("newText is required");
        var replaceAll = arguments?["replaceAll"]?.GetValue<bool>() ?? false;

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var textFragmentAbsorber = new TextFragmentAbsorber(oldText);
        page.Accept(textFragmentAbsorber);

        var fragments = textFragmentAbsorber.TextFragments;
        if (fragments.Count == 0)
            return await Task.FromResult($"Text '{oldText}' not found on page {pageIndex}.");

        int replaceCount = replaceAll ? fragments.Count : 1;
        for (int i = 0; i < replaceCount && i < fragments.Count; i++)
        {
            fragments[i].Text = newText;
        }

        document.Save(outputPath);
        return await Task.FromResult($"Replaced {replaceCount} occurrence(s) of '{oldText}' with '{newText}' on page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> ExtractText(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var includeFontInfo = arguments?["includeFontInfo"]?.GetValue<bool>() ?? false;

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var textAbsorber = new TextAbsorber();
        page.Accept(textAbsorber);

        var sb = new StringBuilder();
        sb.AppendLine($"=== Extracted Text from Page {pageIndex} ===");
        sb.AppendLine();

        if (includeFontInfo)
        {
            var textFragmentAbsorber = new TextFragmentAbsorber();
            page.Accept(textFragmentAbsorber);
            foreach (TextFragment fragment in textFragmentAbsorber.TextFragments)
            {
                sb.AppendLine($"Text: {fragment.Text}");
                sb.AppendLine($"  Font: {fragment.TextState.Font.FontName}, Size: {fragment.TextState.FontSize}");
                sb.AppendLine();
            }
        }
        else
        {
            sb.AppendLine(textAbsorber.Text);
        }

        return await Task.FromResult(sb.ToString());
    }
}

