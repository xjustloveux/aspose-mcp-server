using Aspose.Words;
using Aspose.Words.Drawing;
using WordHeaderFooter = Aspose.Words.HeaderFooter;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word shape operations.
/// </summary>
public static class WordShapeHelper
{
    /// <summary>
    ///     Finds all textboxes in the document, searching in all sections' Body and HeaderFooter.
    ///     This ensures consistent textbox indexing across all operations.
    /// </summary>
    /// <param name="doc">The Word document to search.</param>
    /// <returns>A list of all textbox shapes found in the document.</returns>
    public static List<Shape> FindAllTextboxes(Document doc)
    {
        List<Shape> textboxes = [];
        foreach (var section in doc.Sections.Cast<Section>())
        {
            // Search in main body
            var bodyShapes = section.Body.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                .Where(s => s.ShapeType == ShapeType.TextBox);
            textboxes.AddRange(bodyShapes);

            // Search in headers and footers
            foreach (var header in section.HeadersFooters.Cast<WordHeaderFooter>())
            {
                var headerShapes = header.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                    .Where(s => s.ShapeType == ShapeType.TextBox);
                textboxes.AddRange(headerShapes);
            }
        }

        return textboxes;
    }

    /// <summary>
    ///     Gets all shapes from the document.
    /// </summary>
    /// <param name="doc">The Word document to search.</param>
    /// <returns>A list of all shapes found in the document.</returns>
    public static List<Shape> GetAllShapes(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
    }

    /// <summary>
    ///     Parses a dash style string to DashStyle enum.
    /// </summary>
    /// <param name="borderStyle">The border style string.</param>
    /// <returns>The corresponding DashStyle value.</returns>
    public static DashStyle ParseDashStyle(string borderStyle)
    {
        return borderStyle.ToLower() switch
        {
            "dash" => DashStyle.Dash,
            "dot" => DashStyle.Dot,
            "dashdot" => DashStyle.DashDot,
            "dashdotdot" => DashStyle.LongDashDotDot,
            "rounddot" => DashStyle.ShortDot,
            _ => DashStyle.Solid
        };
    }

    /// <summary>
    ///     Parses a shape type string to ShapeType enum.
    /// </summary>
    /// <param name="shapeType">The shape type string.</param>
    /// <returns>The corresponding ShapeType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the shape type is unknown.</exception>
    public static ShapeType ParseShapeType(string shapeType)
    {
        return shapeType.ToLower() switch
        {
            "rectangle" => ShapeType.Rectangle,
            "ellipse" => ShapeType.Ellipse,
            "roundrectangle" => ShapeType.RoundRectangle,
            "line" => ShapeType.Line,
            _ => throw new ArgumentException($"Unknown shape type: {shapeType}")
        };
    }

    /// <summary>
    ///     Parses alignment string to ParagraphAlignment enum.
    /// </summary>
    /// <param name="alignment">The alignment string.</param>
    /// <returns>The corresponding ParagraphAlignment value.</returns>
    public static ParagraphAlignment ParseAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }
}
