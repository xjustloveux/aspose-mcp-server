using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Words;
using CellsStyle = Aspose.Cells.Style;

namespace AsposeMcpServer.Core;

/// <summary>
///     Unified helper class for font settings across all Aspose tools
///     Organized by tool type: Word, Excel, PowerPoint, PDF
/// </summary>
public static class FontHelper
{
    #region Word Font Settings

    /// <summary>
    ///     Word-specific font helper methods
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
        public static void ApplyFontSettings(
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
            // Apply font names with priority logic
            if (!string.IsNullOrEmpty(fontNameAscii))
                run.Font.NameAscii = fontNameAscii;

            if (!string.IsNullOrEmpty(fontNameFarEast))
                run.Font.NameFarEast = fontNameFarEast;

            if (!string.IsNullOrEmpty(fontName))
            {
                if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                {
                    run.Font.Name = fontName;
                }
                else
                {
                    // If fontNameAscii or fontNameFarEast is set, use fontName as fallback
                    if (string.IsNullOrEmpty(fontNameAscii))
                        run.Font.NameAscii = fontName;
                    if (string.IsNullOrEmpty(fontNameFarEast))
                        run.Font.NameFarEast = fontName;
                }
            }

            // Apply font size
            if (fontSize.HasValue)
                run.Font.Size = fontSize.Value;

            // Apply bold/italic
            if (bold.HasValue)
                run.Font.Bold = bold.Value;

            if (italic.HasValue)
                run.Font.Italic = italic.Value;

            // Apply underline
            if (!string.IsNullOrEmpty(underline))
                run.Font.Underline = ParseUnderline(underline);

            // Apply color (ParseColor returns Color.Black on failure, no exception thrown)
            if (!string.IsNullOrEmpty(color))
                run.Font.Color = ColorHelper.ParseColor(color);

            // Apply strikethrough
            if (strikethrough.HasValue)
                run.Font.StrikeThrough = strikethrough.Value;

            // Apply superscript/subscript (mutually exclusive)
            if (superscript.HasValue || subscript.HasValue)
            {
                if (superscript.HasValue && superscript.Value)
                {
                    run.Font.Subscript = false;
                    run.Font.Superscript = true;
                }
                else if (subscript.HasValue && subscript.Value)
                {
                    run.Font.Superscript = false;
                    run.Font.Subscript = true;
                }
                else
                {
                    // Setting to false
                    if (superscript.HasValue && !superscript.Value)
                        run.Font.Superscript = false;
                    if (subscript.HasValue && !subscript.Value)
                        run.Font.Subscript = false;
                }
            }
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
        public static void ApplyFontSettings(
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
            // Apply font names with priority logic
            if (!string.IsNullOrEmpty(fontNameAscii))
                builder.Font.NameAscii = fontNameAscii;

            if (!string.IsNullOrEmpty(fontNameFarEast))
                builder.Font.NameFarEast = fontNameFarEast;

            if (!string.IsNullOrEmpty(fontName))
            {
                if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                {
                    builder.Font.Name = fontName;
                }
                else
                {
                    // If fontNameAscii or fontNameFarEast is set, use fontName as fallback
                    if (string.IsNullOrEmpty(fontNameAscii))
                        builder.Font.NameAscii = fontName;
                    if (string.IsNullOrEmpty(fontNameFarEast))
                        builder.Font.NameFarEast = fontName;
                }
            }

            // Apply font size
            if (fontSize.HasValue)
                builder.Font.Size = fontSize.Value;

            // Apply bold/italic
            if (bold.HasValue)
                builder.Font.Bold = bold.Value;

            if (italic.HasValue)
                builder.Font.Italic = italic.Value;

            // Apply underline
            if (!string.IsNullOrEmpty(underline))
                builder.Font.Underline = ParseUnderline(underline);

            // Apply color (ParseColor returns Color.Black on failure, no exception thrown)
            if (!string.IsNullOrEmpty(color))
                builder.Font.Color = ColorHelper.ParseColor(color);

            // Apply strikethrough
            if (strikethrough.HasValue)
                builder.Font.StrikeThrough = strikethrough.Value;

            // Apply superscript/subscript (mutually exclusive)
            if (superscript.HasValue || subscript.HasValue)
            {
                if (superscript.HasValue && superscript.Value)
                {
                    builder.Font.Subscript = false;
                    builder.Font.Superscript = true;
                }
                else if (subscript.HasValue && subscript.Value)
                {
                    builder.Font.Superscript = false;
                    builder.Font.Subscript = true;
                }
                else
                {
                    // Setting to false
                    if (superscript.HasValue && !superscript.Value)
                        builder.Font.Superscript = false;
                    if (subscript.HasValue && !subscript.Value)
                        builder.Font.Subscript = false;
                }
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
                "none" => Underline.None,
                _ => Underline.None // Default case for unknown values
            };
        }
    }

    #endregion

    #region Excel Font Settings

    /// <summary>
    ///     Excel-specific font helper methods
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
                // ParseColor returns Color.Black on failure, no exception thrown
                style.Font.Color = ColorHelper.ParseColor(fontColor);
        }
    }

    #endregion

    #region PowerPoint Font Settings

    /// <summary>
    ///     PowerPoint-specific font helper methods
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
                // ParseColor returns Color.Black on failure, no exception thrown
                var colorValue = ColorHelper.ParseColor(color);
                portionFormat.FillFormat.FillType = FillType.Solid;
                portionFormat.FillFormat.SolidFillColor.Color = colorValue;
            }
        }
    }

    #endregion

    #region PDF Font Settings

    /// <summary>
    ///     PDF-specific font helper methods
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
                    // Ignore font not found errors (handled by caller if needed)
                }

            if (fontSize.HasValue)
                textState.FontSize = (float)fontSize.Value;
        }
    }

    #endregion
}