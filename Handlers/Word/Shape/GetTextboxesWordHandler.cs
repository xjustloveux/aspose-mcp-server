using System.Text;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for getting textboxes from Word documents.
/// </summary>
public class GetTextboxesWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_textboxes";

    /// <summary>
    ///     Gets all textboxes from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: includeContent
    /// </param>
    /// <returns>Formatted string containing textbox information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetTextboxesParameters(parameters);

        var doc = context.Document;
        var shapes = WordShapeHelper.FindAllTextboxes(doc);

        var result = new StringBuilder();
        result.AppendLine("=== Document Textboxes ===\n");
        result.AppendLine($"Total Textboxes: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("No textboxes found");
            return result.ToString();
        }

        for (var i = 0; i < shapes.Count; i++)
        {
            var textbox = shapes[i];
            result.AppendLine($"[Textbox {i}]");
            result.AppendLine($"Name: {textbox.Name ?? "(No name)"}");
            result.AppendLine($"Width: {textbox.Width} pt");
            result.AppendLine($"Height: {textbox.Height} pt");
            result.AppendLine($"Position: X={textbox.Left}, Y={textbox.Top}");

            if (p.IncludeContent)
            {
                var textboxText = textbox.GetText().Trim();
                if (!string.IsNullOrEmpty(textboxText))
                {
                    result.AppendLine("Content:");
                    result.AppendLine($"  {textboxText.Replace("\n", "\n  ")}");
                }
                else
                {
                    result.AppendLine("Content: (empty)");
                }
            }

            result.AppendLine();
        }

        return result.ToString();
    }

    private static GetTextboxesParameters ExtractGetTextboxesParameters(OperationParameters parameters)
    {
        var includeContent = parameters.GetOptional("includeContent", true);

        return new GetTextboxesParameters(includeContent);
    }

    private sealed record GetTextboxesParameters(bool IncludeContent);
}
