using Aspose.Words;
using Aspose.Words.Fields;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helpers for locating Word field node boundaries so text and field mutations do not corrupt
///     or nest inside existing fields. Within a paragraph a field is a flat node range:
///     FieldStart -> field-code runs -> FieldSeparator -> result runs -> FieldEnd.
/// </summary>
public static class FieldBoundaryHelper
{
    /// <summary>
    ///     Returns the field whose node range (FieldStart..FieldEnd) contains the specified node,
    ///     or null when the node is not inside any field. The outermost containing field is
    ///     returned when fields are nested.
    /// </summary>
    /// <param name="node">The node to test, typically a Run.</param>
    /// <returns>The enclosing field, or null.</returns>
    public static Field? GetEnclosingField(Node node)
    {
        if (node.GetAncestor(NodeType.Paragraph) is not WordParagraph paragraph)
            return null;

        return paragraph.Range.Fields.FirstOrDefault(field => IsNodeWithinFieldRange(node, field));
    }

    /// <summary>
    ///     Moves the builder cursor to immediately after the field's end so a subsequent insertion
    ///     becomes a sibling following the field rather than landing inside its range.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="field">The field to move past.</param>
    public static void MoveToAfterField(DocumentBuilder builder, Field field)
    {
        if (field.End == null)
            return;

        builder.MoveTo(field.End.NextSibling ?? field.End.ParentNode);
    }

    /// <summary>
    ///     Determines whether a node lies within a field's flat node range (inclusive of the
    ///     FieldStart, up to and excluding the FieldEnd).
    /// </summary>
    /// <param name="node">The node to test.</param>
    /// <param name="field">The field whose range is checked.</param>
    /// <returns>True when the node is within the field range.</returns>
    private static bool IsNodeWithinFieldRange(Node node, Field field)
    {
        if (field.Start == null || field.End == null)
            return false;

        for (var current = (Node?)field.Start;
             current != null && current != field.End;
             current = current.NextSibling)
            if (current == node)
                return true;

        return false;
    }
}
