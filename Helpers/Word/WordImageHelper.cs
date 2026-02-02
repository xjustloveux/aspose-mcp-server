using Aspose.Words;
using Aspose.Words.Drawing;
using Section = Aspose.Words.Section;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word image operations.
/// </summary>
public static class WordImageHelper
{
    /// <summary>
    ///     Gets all images from the document or a specific section.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <returns>A list of Shape objects representing images.</returns>
    /// <exception cref="ArgumentException">Thrown when the section index is out of range.</exception>
    public static List<WordShape> GetAllImages(Document doc, int sectionIndex)
    {
        List<WordShape> allImages = [];

        if (sectionIndex == -1)
        {
            foreach (var section in doc.Sections.Cast<Section>())
            {
                var shapes = section.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage)
                    .ToList();
                allImages.AddRange(shapes);
            }
        }
        else
        {
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException(
                    $"Section index {sectionIndex} is out of range (document has {doc.Sections.Count} sections)");

            var section = doc.Sections[sectionIndex];
            allImages = section.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().Where(s => s.HasImage).ToList();
        }

        return allImages;
    }

    /// <summary>
    ///     Converts alignment string to ParagraphAlignment enum.
    /// </summary>
    /// <param name="alignment">The alignment string (left, center, right).</param>
    /// <returns>The corresponding ParagraphAlignment enum value.</returns>
    public static ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }

    /// <summary>
    ///     Converts wrap type string to WrapType enum.
    /// </summary>
    /// <param name="wrapType">The wrap type string (inline, square, tight, through, topandbottom, none).</param>
    /// <returns>The corresponding WrapType enum value.</returns>
    public static WrapType GetWrapType(string wrapType)
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

    /// <summary>
    ///     Converts alignment string to HorizontalAlignment enum for floating images.
    /// </summary>
    /// <param name="alignment">The alignment string (left, center, right).</param>
    /// <returns>The corresponding HorizontalAlignment enum value.</returns>
    public static HorizontalAlignment GetHorizontalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => HorizontalAlignment.Left,
            "center" => HorizontalAlignment.Center,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Left
        };
    }

    /// <summary>
    ///     Converts alignment string to VerticalAlignment enum for floating images.
    /// </summary>
    /// <param name="alignment">The alignment string (top, center, bottom).</param>
    /// <returns>The corresponding VerticalAlignment enum value.</returns>
    public static VerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => VerticalAlignment.Top,
            "center" => VerticalAlignment.Center,
            "bottom" => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Top
        };
    }

    /// <summary>
    ///     Inserts a professional caption with automatic figure numbering.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="caption">The caption text.</param>
    /// <param name="alignment">The caption alignment.</param>
    public static void InsertCaption(DocumentBuilder builder, string caption, string alignment)
    {
        // Use professional Caption style with SEQ field for automatic figure numbering
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
        builder.ParagraphFormat.Alignment = GetAlignment(alignment);
        builder.Write("Figure ");
        builder.InsertField("SEQ Figure \\* ARABIC");
        builder.Write(": " + caption);
        builder.Writeln();
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
    }
}
