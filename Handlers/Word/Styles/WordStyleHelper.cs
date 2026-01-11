using Aspose.Words;
using WordParagraph = Aspose.Words.Paragraph;
using WordStyle = Aspose.Words.Style;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Helper class for Word style operations.
///     Provides shared utilities for style-related handlers.
/// </summary>
public static class WordStyleHelper
{
    /// <summary>
    ///     Applies a style to a single paragraph, handling empty paragraphs specially.
    /// </summary>
    /// <param name="para">The paragraph to apply the style to.</param>
    /// <param name="style">The style to apply.</param>
    /// <param name="styleName">The name of the style.</param>
    public static void ApplyStyleToParagraph(WordParagraph? para, WordStyle style, string styleName)
    {
        if (para == null)
            throw new ArgumentNullException(nameof(para), "Paragraph cannot be null");

        var paraFormat = para.ParagraphFormat;
        var isEmpty = string.IsNullOrWhiteSpace(para.GetText());

        if (isEmpty)
        {
            paraFormat.ClearFormatting();
            try
            {
                if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                    paraFormat.StyleIdentifier = style.StyleIdentifier;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
            }
        }

        paraFormat.Style = style;
        paraFormat.StyleName = styleName;

        if (isEmpty)
        {
            paraFormat.ClearFormatting();
            paraFormat.Style = style;
            paraFormat.StyleName = styleName;
            try
            {
                if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                    paraFormat.StyleIdentifier = style.StyleIdentifier;
            }
            catch
            {
                // Ignore: some styles may not support StyleIdentifier assignment
            }
        }
    }

    /// <summary>
    ///     Copies style properties from source to target style.
    /// </summary>
    /// <param name="sourceStyle">The source style to copy from.</param>
    /// <param name="targetStyle">The target style to copy to.</param>
    public static void CopyStyleProperties(WordStyle sourceStyle, WordStyle targetStyle)
    {
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
            targetStyle.ParagraphFormat.Alignment = sourceStyle.ParagraphFormat.Alignment;
            targetStyle.ParagraphFormat.SpaceBefore = sourceStyle.ParagraphFormat.SpaceBefore;
            targetStyle.ParagraphFormat.SpaceAfter = sourceStyle.ParagraphFormat.SpaceAfter;
            targetStyle.ParagraphFormat.LineSpacing = sourceStyle.ParagraphFormat.LineSpacing;
            targetStyle.ParagraphFormat.LineSpacingRule = sourceStyle.ParagraphFormat.LineSpacingRule;
            targetStyle.ParagraphFormat.LeftIndent = sourceStyle.ParagraphFormat.LeftIndent;
            targetStyle.ParagraphFormat.RightIndent = sourceStyle.ParagraphFormat.RightIndent;
            targetStyle.ParagraphFormat.FirstLineIndent = sourceStyle.ParagraphFormat.FirstLineIndent;
        }
        else if (sourceStyle.Type == StyleType.Table)
        {
            try
            {
                targetStyle.ParagraphFormat.Alignment = sourceStyle.ParagraphFormat.Alignment;
            }
            catch
            {
                // Ignore: table styles may not support ParagraphFormat properties
            }
        }
    }
}
