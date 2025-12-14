using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfTextTool : IAsposeTool
{
    public string Description => @"Manage text in PDF documents. Supports 3 operations: add, edit, extract.

Usage examples:
- Add text: pdf_text(operation='add', path='doc.pdf', pageIndex=1, text='Hello World', x=100, y=100)
- Edit text: pdf_text(operation='edit', path='doc.pdf', pageIndex=1, text='Updated Text')
- Extract text: pdf_text(operation='extract', path='doc.pdf', pageIndex=1, outputPath='output.txt')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add text to page (required params: path, pageIndex, text)
- 'edit': Edit text on page (required params: path, pageIndex, text)
- 'extract': Extract text from page (required params: path, pageIndex, outputPath)",
                @enum = new[] { "add", "edit", "extract" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
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
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");

        return operation.ToLower() switch
        {
            "add" => await AddText(arguments),
            "edit" => await EditText(arguments),
            "extract" => await ExtractText(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds text to a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, text, x, y, optional fontSize, fontName, fontColor, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddText(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex", "pageIndex");
        var text = ArgumentHelper.GetString(arguments, "text", "text");
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

    /// <summary>
    /// Edits text on a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, textIndex, text, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> EditText(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex", "pageIndex");
        var oldText = ArgumentHelper.GetString(arguments, "oldText", "oldText");
        var newText = ArgumentHelper.GetString(arguments, "newText", "newText");
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

    /// <summary>
    /// Extracts text from a PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional pageIndex</param>
    /// <returns>Extracted text as string</returns>
    private async Task<string> ExtractText(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex", "pageIndex");
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

