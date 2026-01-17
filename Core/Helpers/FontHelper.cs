using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Words;
using CellsStyle = Aspose.Cells.Style;
using Font = Aspose.Words.Font;

namespace AsposeMcpServer.Core.Helpers;

/// <summary>
///     Unified helper class for font settings across all Aspose tools
///     Organized by tool type: Word, Excel, PowerPoint, PDF
/// </summary>
public static class FontHelper
{
    /// <summary>
    ///     Word-specific font helper methods.
    /// </summary>
    public static class Word
    {
        /// <summary>
        ///     Applies font settings to a Word Run object
        /// </summary>
        /// <param name="run">Word Run object to apply font settings to</param>
        /// <param name="fontName">Font name (optional)</param>
        /// <param name="fontNameAscii">Font name for ASCII characters (optional)</param>
        /// <param name="fontNameFarEast">Font name for Far East characters (optional)</param>
        /// <param name="fontSize">Font size in points (optional)</param>
        /// <param name="bold">Bold (optional)</param>
        /// <param name="italic">Italic (optional)</param>
        /// <param name="underline">Underline style: "none", "single", "double", "dotted", "dash" (optional)</param>
        /// <param name="color">Font color in various formats (optional)</param>
        /// <param name="strikethrough">Strikethrough (optional)</param>
        /// <param name="superscript">Superscript (optional)</param>
        /// <param name="subscript">Subscript (optional)</param>
        public static void
            ApplyFontSettings( // NOSONAR S107 - Font utility requires all font settings as optional parameters
                Run run,
                string? fontName = null,
                string? fontNameAscii = null,
                string? fontNameFarEast = null,
                double? fontSize = null,
                bool? bold = null,
                bool? italic = null,
                string? underline = null,
                string? color = null,
                bool? strikethrough = null,
                bool? superscript = null,
                bool? subscript = null)
        {
            ApplyFontNames(run.Font, fontName, fontNameAscii, fontNameFarEast);
            ApplyBasicFontSettings(run.Font, fontSize, bold, italic, underline, color, strikethrough);
            ApplySuperSubscript(run.Font, superscript, subscript);
        }

        /// <summary>
        ///     Applies font settings to a Word DocumentBuilder
        /// </summary>
        /// <param name="builder">Word DocumentBuilder to apply font settings to</param>
        /// <param name="fontName">Font name (optional)</param>
        /// <param name="fontNameAscii">Font name for ASCII characters (optional)</param>
        /// <param name="fontNameFarEast">Font name for Far East characters (optional)</param>
        /// <param name="fontSize">Font size in points (optional)</param>
        /// <param name="bold">Bold (optional)</param>
        /// <param name="italic">Italic (optional)</param>
        /// <param name="underline">Underline style: "none", "single", "double", "dotted", "dash" (optional)</param>
        /// <param name="color">Font color in various formats (optional)</param>
        /// <param name="strikethrough">Strikethrough (optional)</param>
        /// <param name="superscript">Superscript (optional)</param>
        /// <param name="subscript">Subscript (optional)</param>
        public static void
            ApplyFontSettings( // NOSONAR S107 - Font utility requires all font settings as optional parameters
                DocumentBuilder builder,
                string? fontName = null,
                string? fontNameAscii = null,
                string? fontNameFarEast = null,
                double? fontSize = null,
                bool? bold = null,
                bool? italic = null,
                string? underline = null,
                string? color = null,
                bool? strikethrough = null,
                bool? superscript = null,
                bool? subscript = null)
        {
            ApplyFontNames(builder.Font, fontName, fontNameAscii, fontNameFarEast);
            ApplyBasicFontSettings(builder.Font, fontSize, bold, italic, underline, color, strikethrough);
            ApplySuperSubscript(builder.Font, superscript, subscript);
        }

        /// <summary>
        ///     Applies font name settings to a Font object.
        /// </summary>
        /// <param name="font">The font object to apply settings to.</param>
        /// <param name="fontName">The general font name.</param>
        /// <param name="fontNameAscii">The font name for ASCII characters.</param>
        /// <param name="fontNameFarEast">The font name for Far East characters.</param>
        private static void ApplyFontNames(Font font, string? fontName, string? fontNameAscii, string? fontNameFarEast)
        {
            if (!string.IsNullOrEmpty(fontNameAscii))
                font.NameAscii = fontNameAscii;

            if (!string.IsNullOrEmpty(fontNameFarEast))
                font.NameFarEast = fontNameFarEast;

            if (string.IsNullOrEmpty(fontName)) return;

            var hasAscii = !string.IsNullOrEmpty(fontNameAscii);
            var hasFarEast = !string.IsNullOrEmpty(fontNameFarEast);

            if (!hasAscii && !hasFarEast)
            {
                font.Name = fontName;
                return;
            }

            if (!hasAscii) font.NameAscii = fontName;
            if (!hasFarEast) font.NameFarEast = fontName;
        }

        /// <summary>
        ///     Applies basic font settings (size, bold, italic, underline, color, strikethrough).
        /// </summary>
        /// <param name="font">The font object to apply settings to.</param>
        /// <param name="fontSize">The font size in points.</param>
        /// <param name="bold">Whether to apply bold.</param>
        /// <param name="italic">Whether to apply italic.</param>
        /// <param name="underline">The underline style.</param>
        /// <param name="color">The font color.</param>
        /// <param name="strikethrough">Whether to apply strikethrough.</param>
        private static void ApplyBasicFontSettings(Font font, double? fontSize, bool? bold, bool? italic,
            string? underline, string? color, bool? strikethrough)
        {
            if (fontSize.HasValue) font.Size = fontSize.Value;
            if (bold.HasValue) font.Bold = bold.Value;
            if (italic.HasValue) font.Italic = italic.Value;
            if (!string.IsNullOrEmpty(underline)) font.Underline = ParseUnderline(underline);
            if (!string.IsNullOrEmpty(color)) font.Color = ColorHelper.ParseColor(color);
            if (strikethrough.HasValue) font.StrikeThrough = strikethrough.Value;
        }

        /// <summary>
        ///     Applies superscript and subscript settings to a Font object.
        /// </summary>
        /// <param name="font">The font object to apply settings to.</param>
        /// <param name="superscript">Whether to apply superscript.</param>
        /// <param name="subscript">Whether to apply subscript.</param>
        private static void ApplySuperSubscript(Font font, bool? superscript, bool? subscript)
        {
            if (!superscript.HasValue && !subscript.HasValue) return;

            if (superscript is true)
            {
                font.Subscript = false;
                font.Superscript = true;
            }
            else if (subscript is true)
            {
                font.Superscript = false;
                font.Subscript = true;
            }
            else
            {
                if (superscript is false) font.Superscript = false;
                if (subscript is false) font.Subscript = false;
            }
        }

        /// <summary>
        ///     Parses underline string to Underline enum
        /// </summary>
        /// <param name="underline">Underline style string</param>
        /// <returns>Underline enum value</returns>
        public static Underline ParseUnderline(string? underline)
        {
            if (string.IsNullOrEmpty(underline))
                return Underline.None;

            return underline.ToLower() switch
            {
                "single" => Underline.Single,
                "double" => Underline.Double,
                "dotted" => Underline.Dotted,
                "dash" => Underline.Dash,
                _ => Underline.None
            };
        }
    }

    /// <summary>
    ///     Excel-specific font helper methods.
    /// </summary>
    public static class Excel
    {
        /// <summary>
        ///     Applies font settings to an Excel Style object
        /// </summary>
        /// <param name="style">Excel Style object to apply font settings to</param>
        /// <param name="fontName">Font name (optional)</param>
        /// <param name="fontSize">Font size (optional)</param>
        /// <param name="bold">Bold (optional)</param>
        /// <param name="italic">Italic (optional)</param>
        /// <param name="fontColor">Font color in hex format (optional)</param>
        public static void ApplyFontSettings(
            CellsStyle style,
            string? fontName = null,
            int? fontSize = null,
            bool? bold = null,
            bool? italic = null,
            string? fontColor = null)
        {
            if (!string.IsNullOrEmpty(fontName))
                style.Font.Name = fontName;

            if (fontSize.HasValue)
                style.Font.Size = fontSize.Value;

            if (bold.HasValue)
                style.Font.IsBold = bold.Value;

            if (italic.HasValue)
                style.Font.IsItalic = italic.Value;

            if (!string.IsNullOrWhiteSpace(fontColor))
                style.Font.Color = ColorHelper.ParseColor(fontColor);
        }
    }

    /// <summary>
    ///     PowerPoint-specific font helper methods.
    /// </summary>
    public static class Ppt
    {
        /// <summary>
        ///     Applies font settings to a PowerPoint PortionFormat object
        /// </summary>
        /// <param name="portionFormat">PowerPoint PortionFormat object to apply font settings to</param>
        /// <param name="fontName">Font name (optional)</param>
        /// <param name="fontSize">Font size (optional)</param>
        /// <param name="bold">Bold (optional)</param>
        /// <param name="italic">Italic (optional)</param>
        /// <param name="color">Font color in hex format (optional)</param>
        public static void ApplyFontSettings(
            IPortionFormat portionFormat,
            string? fontName = null,
            double? fontSize = null,
            bool? bold = null,
            bool? italic = null,
            string? color = null)
        {
            if (!string.IsNullOrWhiteSpace(fontName))
                portionFormat.LatinFont = new FontData(fontName);

            if (fontSize.HasValue)
                portionFormat.FontHeight = (float)fontSize.Value;

            if (bold.HasValue)
                portionFormat.FontBold = bold.Value ? NullableBool.True : NullableBool.False;

            if (italic.HasValue)
                portionFormat.FontItalic = italic.Value ? NullableBool.True : NullableBool.False;

            if (!string.IsNullOrWhiteSpace(color))
            {
                var colorValue = ColorHelper.ParseColor(color);
                portionFormat.FillFormat.FillType = FillType.Solid;
                portionFormat.FillFormat.SolidFillColor.Color = colorValue;
            }
        }
    }

    /// <summary>
    ///     PDF-specific font helper methods.
    /// </summary>
    public static class Pdf
    {
        /// <summary>
        ///     Applies font settings to a PDF TextState object
        /// </summary>
        /// <param name="textState">PDF TextState object to apply font settings to</param>
        /// <param name="fontName">Font name (optional)</param>
        /// <param name="fontSize">Font size (optional)</param>
        public static void ApplyFontSettings(
            TextState textState,
            string? fontName = null,
            double? fontSize = null)
        {
            if (!string.IsNullOrWhiteSpace(fontName))
                try
                {
                    textState.Font = FontRepository.FindFont(fontName);
                }
                catch
                {
                    // Ignore font not found errors, use default font
                }

            if (fontSize.HasValue)
                textState.FontSize = (float)fontSize.Value;
        }
    }
}
