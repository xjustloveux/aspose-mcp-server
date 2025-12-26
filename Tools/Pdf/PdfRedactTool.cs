using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using Color = System.Drawing.Color;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for redacting (blacking out) text or areas on PDF pages
/// </summary>
public class PdfRedactTool : IAsposeTool
{
    public string Description => @"Redact (black out) text or area on PDF page.

Usage examples:
- Redact area: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50)
- Redact with color: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50, fillColor='255,0,0')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path (required)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based)"
            },
            x = new
            {
                type = "number",
                description = "X position of redaction area"
            },
            y = new
            {
                type = "number",
                description = "Y position of redaction area"
            },
            width = new
            {
                type = "number",
                description = "Width of redaction area"
            },
            height = new
            {
                type = "number",
                description = "Height of redaction area"
            },
            fillColor = new
            {
                type = "string",
                description = "Fill color (optional, default: black, format: 'R,G,B' or color name)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "path", "pageIndex", "x", "y", "width", "height" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var x = ArgumentHelper.GetDouble(arguments, "x");
            var y = ArgumentHelper.GetDouble(arguments, "y");
            var width = ArgumentHelper.GetDouble(arguments, "width");
            var height = ArgumentHelper.GetDouble(arguments, "height");
            var fillColor = ArgumentHelper.GetStringNullable(arguments, "fillColor");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            var rect = new Rectangle(x, y, x + width, y + height);

            var redactionAnnotation = new RedactionAnnotation(page, rect);

            // Set fill color
            if (!string.IsNullOrEmpty(fillColor))
            {
                // Parse color with fallback to black if parsing fails
                var systemColor = ColorHelper.ParseColor(fillColor, Color.Black);
                redactionAnnotation.FillColor = ColorHelper.ToPdfColor(systemColor);
            }
            else
            {
                redactionAnnotation.FillColor = Aspose.Pdf.Color.Black;
            }

            page.Annotations.Add(redactionAnnotation);
            document.Save(outputPath);

            return $"Redaction applied to page {pageIndex}. Output: {outputPath}";
        });
    }
}