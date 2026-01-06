using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for header and footer operations in Word documents
///     Merges: WordSetHeaderTextTool, WordSetFooterTextTool, WordSetHeaderImageTool, WordSetFooterImageTool,
///     WordSetHeaderLineTool, WordSetFooterLineTool, WordSetHeaderTabStopsTool, WordSetFooterTabStopsTool,
///     WordSetHeaderFooterTool, WordGetHeadersFootersTool
/// </summary>
[McpServerToolType]
public class WordHeaderFooterTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordHeaderFooterTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordHeaderFooterTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word header/footer operation (set_header_text, set_footer_text, set_header_image, set_footer_image,
    ///     set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: set_header_text, set_footer_text, set_header_image, set_footer_image,
    ///     set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="headerLeft">Header left section text (for set_header_text).</param>
    /// <param name="headerCenter">Header center section text (for set_header_text).</param>
    /// <param name="headerRight">Header right section text (for set_header_text).</param>
    /// <param name="footerLeft">Footer left section text (for set_footer_text).</param>
    /// <param name="footerCenter">Footer center section text (for set_footer_text).</param>
    /// <param name="footerRight">Footer right section text (for set_footer_text).</param>
    /// <param name="imagePath">Path to image file (for set_header_image/set_footer_image).</param>
    /// <param name="alignment">Image alignment: left, center, right (for image operations).</param>
    /// <param name="imageWidth">Image width in points (for image operations).</param>
    /// <param name="imageHeight">Image height in points (for image operations).</param>
    /// <param name="lineStyle">Line style: single, double, thick (for line operations).</param>
    /// <param name="lineWidth">Line width in points (for line operations).</param>
    /// <param name="tabStops">Tab stops array (for tab operations).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontNameAscii">Font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">Font name for Far East characters.</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="headerFooterType">Header/footer type: Primary, FirstPage, EvenPage.</param>
    /// <param name="isFloating">Make image floating instead of inline.</param>
    /// <param name="autoTabStops">Automatically add tab stops when using left/center/right text.</param>
    /// <param name="clearExisting">Clear existing content before setting new content.</param>
    /// <param name="clearTextOnly">Only clear text content, preserve images and shapes.</param>
    /// <param name="removeExisting">Remove existing images before adding new one.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_header_footer")]
    [Description(
        @"Manage headers and footers in Word documents. Supports 10 operations: set_header_text, set_footer_text, set_header_image, set_footer_image, set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get.

Usage examples:
- Set header text: word_header_footer(operation='set_header_text', path='doc.docx', headerLeft='Left', headerCenter='Center', headerRight='Right')
- Set footer text: word_header_footer(operation='set_footer_text', path='doc.docx', footerLeft='Page', footerCenter='', footerRight='{PAGE}')
- Set header image: word_header_footer(operation='set_header_image', path='doc.docx', imagePath='logo.png')
- Get headers/footers: word_header_footer(operation='get', path='doc.docx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'set_header_text': Set header text (required params: path)
- 'set_footer_text': Set footer text (required params: path)
- 'set_header_image': Set header image (required params: path, imagePath)
- 'set_footer_image': Set footer image (required params: path, imagePath)
- 'set_header_line': Set header line (required params: path)
- 'set_footer_line': Set footer line (required params: path)
- 'set_header_tabs': Set header tab stops (required params: path)
- 'set_footer_tabs': Set footer tab stops (required params: path)
- 'set_header_footer': Set header and footer together (required params: path)
- 'get': Get headers and footers info (required params: path)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Header left section text (optional, for set_header_text operation)")]
        string? headerLeft = null,
        [Description("Header center section text (optional, for set_header_text operation)")]
        string? headerCenter = null,
        [Description("Header right section text (optional, for set_header_text operation)")]
        string? headerRight = null,
        [Description("Footer left section text (optional, for set_footer_text operation)")]
        string? footerLeft = null,
        [Description("Footer center section text (optional, for set_footer_text operation)")]
        string? footerCenter = null,
        [Description("Footer right section text (optional, for set_footer_text operation)")]
        string? footerRight = null,
        [Description("Path to image file (required for set_header_image/set_footer_image operations)")]
        string? imagePath = null,
        [Description("Image alignment: left, center, right (optional, default: left, for image operations)")]
        string alignment = "left",
        [Description("Image width in points (optional, default: 50, for image operations)")]
        double? imageWidth = null,
        [Description(
            "Image height in points (optional, maintains aspect ratio if not specified, for image operations)")]
        double? imageHeight = null,
        [Description("Line style: single, double, thick (optional, for line operations)")]
        string lineStyle = "single",
        [Description("Line width in points (optional, for line operations)")]
        double? lineWidth = null,
        [Description("Tab stops (optional, for tab operations)")]
        JsonArray? tabStops = null,
        [Description("Font name (optional)")] string? fontName = null,
        [Description("Font name for ASCII characters (English, optional)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (Chinese/Japanese/Korean, optional)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (optional)")]
        double? fontSize = null,
        [Description("Section index (0-based, optional, default: 0, use -1 to apply to all sections)")]
        int sectionIndex = 0,
        [Description(
            "Header/Footer type: primary (default), firstPage, evenPages. Use firstPage for different first page, evenPages for odd/even page layouts.")]
        string headerFooterType = "primary",
        [Description(
            "Make image floating instead of inline (optional, default: false, for image operations). Floating images can be precisely positioned.")]
        bool isFloating = false,
        [Description(
            "Automatically add center and right tab stops when using left/center/right text (optional, default: true, for text operations)")]
        bool autoTabStops = true,
        [Description("Clear existing content before setting new content (optional, default: true)")]
        bool clearExisting = true,
        [Description(
            "Only clear text content, preserve images and shapes (optional, default: false, for text operations)")]
        bool clearTextOnly = false,
        [Description("Remove existing images before adding new one (optional, default: true, for image operations)")]
        bool removeExisting = true)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "set_header_text" => SetHeaderText(ctx, outputPath, headerLeft, headerCenter, headerRight, fontName,
                fontNameAscii, fontNameFarEast, fontSize, sectionIndex, headerFooterType, autoTabStops, clearExisting,
                clearTextOnly),
            "set_footer_text" => SetFooterText(ctx, outputPath, footerLeft, footerCenter, footerRight, fontName,
                fontNameAscii, fontNameFarEast, fontSize, sectionIndex, headerFooterType, autoTabStops, clearExisting,
                clearTextOnly),
            "set_header_image" => SetHeaderImage(ctx, outputPath, imagePath!, alignment, imageWidth, imageHeight,
                sectionIndex, headerFooterType, isFloating, removeExisting),
            "set_footer_image" => SetFooterImage(ctx, outputPath, imagePath!, alignment, imageWidth, imageHeight,
                sectionIndex, headerFooterType, isFloating, removeExisting),
            "set_header_line" => SetHeaderLine(ctx, outputPath, lineStyle, lineWidth, sectionIndex, headerFooterType),
            "set_footer_line" => SetFooterLine(ctx, outputPath, lineStyle, lineWidth, sectionIndex, headerFooterType),
            "set_header_tabs" => SetHeaderTabStops(ctx, outputPath, tabStops, sectionIndex, headerFooterType),
            "set_footer_tabs" => SetFooterTabStops(ctx, outputPath, tabStops, sectionIndex, headerFooterType),
            "set_header_footer" => SetHeaderFooter(ctx, outputPath, headerLeft, headerCenter, headerRight, footerLeft,
                footerCenter, footerRight, fontName, fontNameAscii, fontNameFarEast, fontSize, sectionIndex,
                headerFooterType, autoTabStops, clearExisting, clearTextOnly),
            "get" => GetHeadersFooters(ctx, sectionIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets header text for the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="headerLeft">The left section text of the header.</param>
    /// <param name="headerCenter">The center section text of the header.</param>
    /// <param name="headerRight">The right section text of the header.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <param name="autoTabStops">Whether to automatically add tab stops.</param>
    /// <param name="clearExisting">Whether to clear existing content.</param>
    /// <param name="clearTextOnly">Whether to clear only text content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetHeaderText(DocumentContext<Document> ctx, string? outputPath, string? headerLeft,
        string? headerCenter, string? headerRight,
        string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize,
        int sectionIndex, string headerFooterTypeStr, bool autoTabStops, bool clearExisting, bool clearTextOnly)
    {
        var doc = ctx.Document;

        var hasContent = !string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) ||
                         !string.IsNullOrEmpty(headerRight);
        if (!hasContent)
            return "Warning: No header text content provided";

        // Get the appropriate HeaderFooterType
        var hfType = GetHeaderFooterType(headerFooterTypeStr, true);

        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.HeaderFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.HeaderEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var header = section.HeadersFooters[hfType];
            if (header == null)
            {
                header = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(header);
            }
            else if (clearExisting)
            {
                if (clearTextOnly)
                    ClearTextOnly(header);
                else
                    header.RemoveAllChildren();
            }

            if (hasContent)
            {
                if (!clearTextOnly)
                    header.RemoveAllChildren();

                var headerPara = new Paragraph(doc);
                header.AppendChild(headerPara);

                if (autoTabStops && (!string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight)))
                {
                    var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin -
                                    section.PageSetup.RightMargin;
                    headerPara.ParagraphFormat.TabStops.Clear();
                    // Center tab stop at middle of page
                    headerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth / 2, TabAlignment.Center,
                        TabLeader.None));
                    // Right tab stop at right margin
                    headerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth, TabAlignment.Right,
                        TabLeader.None));
                }

                var builder = new DocumentBuilder(doc);
                builder.MoveTo(headerPara);

                if (!string.IsNullOrEmpty(headerLeft))
                    InsertTextOrField(builder, headerLeft, fontName, fontNameAscii, fontNameFarEast, fontSize);

                if (!string.IsNullOrEmpty(headerCenter))
                {
                    builder.Write("\t");
                    InsertTextOrField(builder, headerCenter, fontName, fontNameAscii, fontNameFarEast, fontSize);
                }

                if (!string.IsNullOrEmpty(headerRight))
                {
                    builder.Write("\t");
                    InsertTextOrField(builder, headerRight, fontName, fontNameAscii, fontNameFarEast, fontSize);
                }
            }
        }

        ctx.Save(outputPath);

        List<string> contentParts = [];
        if (!string.IsNullOrEmpty(headerLeft)) contentParts.Add("left");
        if (!string.IsNullOrEmpty(headerCenter)) contentParts.Add("center");
        if (!string.IsNullOrEmpty(headerRight)) contentParts.Add("right");

        var contentDesc = string.Join(", ", contentParts);
        var sectionsDesc = sectionIndex == -1 ? "all sections" : $"section {sectionIndex}";

        var result = $"Header text set successfully ({contentDesc}) in {sectionsDesc}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets footer text for the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="footerLeft">The left section text of the footer.</param>
    /// <param name="footerCenter">The center section text of the footer.</param>
    /// <param name="footerRight">The right section text of the footer.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <param name="autoTabStops">Whether to automatically add tab stops.</param>
    /// <param name="clearExisting">Whether to clear existing content.</param>
    /// <param name="clearTextOnly">Whether to clear only text content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetFooterText(DocumentContext<Document> ctx, string? outputPath, string? footerLeft,
        string? footerCenter, string? footerRight,
        string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize,
        int sectionIndex, string headerFooterTypeStr, bool autoTabStops, bool clearExisting, bool clearTextOnly)
    {
        var doc = ctx.Document;

        var hasContent = !string.IsNullOrEmpty(footerLeft) || !string.IsNullOrEmpty(footerCenter) ||
                         !string.IsNullOrEmpty(footerRight);
        if (!hasContent)
            return "Warning: No footer text content provided";

        // Get the appropriate HeaderFooterType
        var hfType = GetHeaderFooterType(headerFooterTypeStr, false);

        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.FooterFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.FooterEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var footer = section.HeadersFooters[hfType];
            if (footer == null)
            {
                footer = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(footer);
            }
            else if (clearExisting)
            {
                if (clearTextOnly)
                    ClearTextOnly(footer);
                else
                    footer.RemoveAllChildren();
            }

            if (hasContent)
            {
                if (!clearTextOnly)
                    footer.RemoveAllChildren();

                var footerPara = new Paragraph(doc);
                footer.AppendChild(footerPara);

                if (autoTabStops && (!string.IsNullOrEmpty(footerCenter) || !string.IsNullOrEmpty(footerRight)))
                {
                    var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin -
                                    section.PageSetup.RightMargin;
                    footerPara.ParagraphFormat.TabStops.Clear();
                    // Center tab stop at middle of page
                    footerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth / 2, TabAlignment.Center,
                        TabLeader.None));
                    // Right tab stop at right margin
                    footerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth, TabAlignment.Right,
                        TabLeader.None));
                }

                var builder = new DocumentBuilder(doc);
                builder.MoveTo(footerPara);

                if (!string.IsNullOrEmpty(footerLeft))
                    InsertTextOrField(builder, footerLeft, fontName, fontNameAscii, fontNameFarEast, fontSize);

                if (!string.IsNullOrEmpty(footerCenter))
                {
                    builder.Write("\t");
                    InsertTextOrField(builder, footerCenter, fontName, fontNameAscii, fontNameFarEast, fontSize);
                }

                if (!string.IsNullOrEmpty(footerRight))
                {
                    builder.Write("\t");
                    InsertTextOrField(builder, footerRight, fontName, fontNameAscii, fontNameFarEast, fontSize);
                }
            }
        }

        ctx.Save(outputPath);

        List<string> contentParts = [];
        if (!string.IsNullOrEmpty(footerLeft)) contentParts.Add("left");
        if (!string.IsNullOrEmpty(footerCenter)) contentParts.Add("center");
        if (!string.IsNullOrEmpty(footerRight)) contentParts.Add("right");

        var contentDesc = string.Join(", ", contentParts);
        var sectionsDesc = sectionIndex == -1 ? "all sections" : $"section {sectionIndex}";

        var result = $"Footer text set successfully ({contentDesc}) in {sectionsDesc}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets header image for the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="alignment">The image alignment (left, center, right).</param>
    /// <param name="imageWidth">The image width in points.</param>
    /// <param name="imageHeight">The image height in points.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <param name="isFloating">Whether the image is floating.</param>
    /// <param name="removeExisting">Whether to remove existing images.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    private static string SetHeaderImage(DocumentContext<Document> ctx, string? outputPath, string imagePath,
        string alignment, double? imageWidth, double? imageHeight,
        int sectionIndex, string headerFooterTypeStr, bool isFloating, bool removeExisting)
    {
        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = ctx.Document;
        var hfType = GetHeaderFooterType(headerFooterTypeStr, true);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.HeaderFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.HeaderEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var header = section.HeadersFooters[hfType];
            if (header == null)
            {
                header = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(header);
            }

            if (removeExisting)
            {
                var existingShapes = header.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                    .Where(s => s.HasImage).ToList();
                foreach (var existingShape in existingShapes) existingShape.Remove();
            }

            // Create a new paragraph for the image
            var headerPara = new Paragraph(doc);
            header.AppendChild(headerPara);

            var builder = new DocumentBuilder(doc);
            builder.MoveTo(headerPara);

            // Set paragraph alignment before inserting image (for inline images)
            var paraAlignment = alignment.ToLower() switch
            {
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                _ => ParagraphAlignment.Left
            };
            builder.ParagraphFormat.Alignment = paraAlignment;

            // Insert image
            var shape = builder.InsertImage(imagePath);
            if (imageWidth.HasValue) shape.Width = imageWidth.Value;
            if (imageHeight.HasValue) shape.Height = imageHeight.Value;

            if (isFloating)
            {
                shape.WrapType = WrapType.Square;
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.TopMargin;

                // Position based on alignment
                var pageWidth = section.PageSetup.PageWidth;
                var leftMargin = section.PageSetup.LeftMargin;
                var rightMargin = section.PageSetup.RightMargin;

                switch (alignment.ToLower())
                {
                    case "center":
                        shape.Left = (pageWidth - shape.Width) / 2;
                        break;
                    case "right":
                        shape.Left = pageWidth - rightMargin - shape.Width;
                        break;
                    default: // left
                        shape.Left = leftMargin;
                        break;
                }

                shape.Top = 0; // Top of margin area
            }
            else
            {
                // Ensure the paragraph containing the shape has correct alignment
                headerPara.ParagraphFormat.Alignment = paraAlignment;
            }
        }

        ctx.Save(outputPath);
        var floatingDesc = isFloating ? " (floating)" : "";
        var result = $"Header image set{floatingDesc}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets footer image for the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="alignment">The image alignment (left, center, right).</param>
    /// <param name="imageWidth">The image width in points.</param>
    /// <param name="imageHeight">The image height in points.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <param name="isFloating">Whether the image is floating.</param>
    /// <param name="removeExisting">Whether to remove existing images.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    private static string SetFooterImage(DocumentContext<Document> ctx, string? outputPath, string imagePath,
        string alignment, double? imageWidth, double? imageHeight,
        int sectionIndex, string headerFooterTypeStr, bool isFloating, bool removeExisting)
    {
        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = ctx.Document;
        var hfType = GetHeaderFooterType(headerFooterTypeStr, false);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.FooterFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.FooterEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var footer = section.HeadersFooters[hfType];
            if (footer == null)
            {
                footer = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(footer);
            }

            if (removeExisting)
            {
                var existingShapes = footer.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                    .Where(s => s.HasImage).ToList();
                foreach (var existingShape in existingShapes) existingShape.Remove();
            }

            // Create a new paragraph for the image
            var footerPara = new Paragraph(doc);
            footer.AppendChild(footerPara);

            var builder = new DocumentBuilder(doc);
            builder.MoveTo(footerPara);

            // Set paragraph alignment before inserting image (for inline images)
            var paraAlignment = alignment.ToLower() switch
            {
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                _ => ParagraphAlignment.Left
            };
            builder.ParagraphFormat.Alignment = paraAlignment;

            // Insert image
            var shape = builder.InsertImage(imagePath);
            if (imageWidth.HasValue) shape.Width = imageWidth.Value;
            if (imageHeight.HasValue) shape.Height = imageHeight.Value;

            if (isFloating)
            {
                shape.WrapType = WrapType.Square;
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.BottomMargin;

                // Position based on alignment
                var pageWidth = section.PageSetup.PageWidth;
                var leftMargin = section.PageSetup.LeftMargin;
                var rightMargin = section.PageSetup.RightMargin;

                switch (alignment.ToLower())
                {
                    case "center":
                        shape.Left = (pageWidth - shape.Width) / 2;
                        break;
                    case "right":
                        shape.Left = pageWidth - rightMargin - shape.Width;
                        break;
                    default: // left
                        shape.Left = leftMargin;
                        break;
                }

                shape.Top = 0; // Top of margin area (bottom margin)
            }
            else
            {
                // Ensure the paragraph containing the shape has correct alignment
                footerPara.ParagraphFormat.Alignment = paraAlignment;
            }
        }

        ctx.Save(outputPath);
        var floatingDesc = isFloating ? " (floating)" : "";
        var result = $"Footer image set{floatingDesc}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets header line (border) for the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="lineStyle">The line style (single, double, thick).</param>
    /// <param name="lineWidth">The line width in points.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetHeaderLine(DocumentContext<Document> ctx, string? outputPath, string lineStyle,
        double? lineWidth,
        int sectionIndex, string headerFooterTypeStr)
    {
        var doc = ctx.Document;
        var hfType = GetHeaderFooterType(headerFooterTypeStr, true);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.HeaderFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.HeaderEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var header = section.HeadersFooters[hfType];
            if (header == null)
            {
                header = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(header);
            }

            var para = new Paragraph(doc);
            para.ParagraphFormat.Borders.Bottom.LineStyle = lineStyle.ToLower() switch
            {
                "double" => LineStyle.Double,
                "thick" => LineStyle.Thick,
                _ => LineStyle.Single
            };

            if (lineWidth.HasValue) para.ParagraphFormat.Borders.Bottom.LineWidth = lineWidth.Value;

            header.AppendChild(para);
        }

        ctx.Save(outputPath);
        var result = "Header line set\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets footer line (border) for the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="lineStyle">The line style (single, double, thick).</param>
    /// <param name="lineWidth">The line width in points.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetFooterLine(DocumentContext<Document> ctx, string? outputPath, string lineStyle,
        double? lineWidth,
        int sectionIndex, string headerFooterTypeStr)
    {
        var doc = ctx.Document;
        var hfType = GetHeaderFooterType(headerFooterTypeStr, false);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.FooterFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.FooterEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var footer = section.HeadersFooters[hfType];
            if (footer == null)
            {
                footer = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(footer);
            }

            var para = new Paragraph(doc);
            para.ParagraphFormat.Borders.Top.LineStyle = lineStyle.ToLower() switch
            {
                "double" => LineStyle.Double,
                "thick" => LineStyle.Thick,
                _ => LineStyle.Single
            };

            if (lineWidth.HasValue) para.ParagraphFormat.Borders.Top.LineWidth = lineWidth.Value;

            footer.AppendChild(para);
        }

        ctx.Save(outputPath);
        var result = "Footer line set\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets header tab stops.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tabStops">The tab stops configuration as JSON array.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetHeaderTabStops(DocumentContext<Document> ctx, string? outputPath, JsonArray? tabStops,
        int sectionIndex, string headerFooterTypeStr)
    {
        var doc = ctx.Document;
        var hfType = GetHeaderFooterType(headerFooterTypeStr, true);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.HeaderFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.HeaderEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var header = section.HeadersFooters[hfType];
            if (header == null)
            {
                header = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(header);
            }

            if (tabStops is { Count: > 0 })
            {
                var para = header.FirstParagraph ?? new Paragraph(doc);
                para.ParagraphFormat.TabStops.Clear();

                foreach (var tabStopJson in tabStops)
                {
                    var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
                    var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "left";
                    var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "none";

                    var tabAlignment = alignmentStr.ToLower() switch
                    {
                        "center" => TabAlignment.Center,
                        "right" => TabAlignment.Right,
                        "decimal" => TabAlignment.Decimal,
                        "bar" => TabAlignment.Bar,
                        _ => TabAlignment.Left
                    };

                    var tabLeader = leaderStr.ToLower() switch
                    {
                        "dots" => TabLeader.Dots,
                        "dashes" => TabLeader.Dashes,
                        "line" => TabLeader.Line,
                        _ => TabLeader.None
                    };

                    para.ParagraphFormat.TabStops.Add(new TabStop(position, tabAlignment, tabLeader));
                }

                if (header.FirstParagraph == null) header.AppendChild(para);
            }
        }

        ctx.Save(outputPath);
        var result = "Header tab stops set\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets footer tab stops.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tabStops">The tab stops configuration as JSON array.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetFooterTabStops(DocumentContext<Document> ctx, string? outputPath, JsonArray? tabStops,
        int sectionIndex, string headerFooterTypeStr)
    {
        var doc = ctx.Document;
        var hfType = GetHeaderFooterType(headerFooterTypeStr, false);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            // Enable different first page or odd/even if needed
            if (hfType == HeaderFooterType.FooterFirst)
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            else if (hfType == HeaderFooterType.FooterEven)
                section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            var footer = section.HeadersFooters[hfType];
            if (footer == null)
            {
                footer = new HeaderFooter(doc, hfType);
                section.HeadersFooters.Add(footer);
            }

            if (tabStops is { Count: > 0 })
            {
                var para = footer.FirstParagraph ?? new Paragraph(doc);
                para.ParagraphFormat.TabStops.Clear();

                foreach (var tabStopJson in tabStops)
                {
                    var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
                    var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "left";
                    var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "none";

                    var tabAlignment = alignmentStr.ToLower() switch
                    {
                        "center" => TabAlignment.Center,
                        "right" => TabAlignment.Right,
                        "decimal" => TabAlignment.Decimal,
                        "bar" => TabAlignment.Bar,
                        _ => TabAlignment.Left
                    };

                    var tabLeader = leaderStr.ToLower() switch
                    {
                        "dots" => TabLeader.Dots,
                        "dashes" => TabLeader.Dashes,
                        "line" => TabLeader.Line,
                        _ => TabLeader.None
                    };

                    para.ParagraphFormat.TabStops.Add(new TabStop(position, tabAlignment, tabLeader));
                }

                if (footer.FirstParagraph == null) footer.AppendChild(para);
            }
        }

        ctx.Save(outputPath);
        var result = "Footer tab stops set\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets header and footer together.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="headerLeft">The left section text of the header.</param>
    /// <param name="headerCenter">The center section text of the header.</param>
    /// <param name="headerRight">The right section text of the header.</param>
    /// <param name="footerLeft">The left section text of the footer.</param>
    /// <param name="footerCenter">The center section text of the footer.</param>
    /// <param name="footerRight">The right section text of the footer.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="headerFooterTypeStr">The header/footer type (primary, firstPage, evenPages).</param>
    /// <param name="autoTabStops">Whether to automatically add tab stops.</param>
    /// <param name="clearExisting">Whether to clear existing content.</param>
    /// <param name="clearTextOnly">Whether to clear only text content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetHeaderFooter(DocumentContext<Document> ctx, string? outputPath,
        string? headerLeft, string? headerCenter, string? headerRight,
        string? footerLeft, string? footerCenter, string? footerRight,
        string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize,
        int sectionIndex, string headerFooterTypeStr, bool autoTabStops, bool clearExisting, bool clearTextOnly)
    {
        // Combined operation to set both header and footer
        SetHeaderText(ctx, null, headerLeft, headerCenter, headerRight, fontName, fontNameAscii, fontNameFarEast,
            fontSize, sectionIndex, headerFooterTypeStr, autoTabStops, clearExisting, clearTextOnly);
        SetFooterText(ctx, null, footerLeft, footerCenter, footerRight, fontName, fontNameAscii, fontNameFarEast,
            fontSize, sectionIndex, headerFooterTypeStr, autoTabStops, clearExisting, clearTextOnly);

        ctx.Save(outputPath);
        var result = "Header and footer set\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets all headers and footers from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <returns>A JSON string containing header and footer information.</returns>
    /// <exception cref="ArgumentException">Thrown when the section index is out of range.</exception>
    private static string GetHeadersFooters(DocumentContext<Document> ctx, int? sectionIndex)
    {
        var doc = ctx.Document;
        doc.UpdateFields();

        var sections = sectionIndex.HasValue && sectionIndex.Value != -1
            ? [doc.Sections[sectionIndex.Value]]
            : doc.Sections.Cast<Section>().ToArray();

        if (sectionIndex.HasValue && sectionIndex.Value != -1 &&
            (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
            throw new ArgumentException(
                $"Section index {sectionIndex.Value} is out of range (document has {doc.Sections.Count} sections)");

        List<object> sectionsList = [];

        for (var i = 0; i < sections.Length; i++)
        {
            var section = sections[i];
            var actualIndex = sectionIndex.HasValue && sectionIndex.Value != -1 ? sectionIndex.Value : i;

            var headerTypes = new[]
            {
                (HeaderFooterType.HeaderPrimary, "primary"),
                (HeaderFooterType.HeaderFirst, "firstPage"),
                (HeaderFooterType.HeaderEven, "evenPage")
            };

            var headers = new Dictionary<string, string?>();
            foreach (var (type, name) in headerTypes)
            {
                var header = section.HeadersFooters[type];
                if (header != null)
                {
                    var headerText = header.ToString(SaveFormat.Text).Trim();
                    if (!string.IsNullOrEmpty(headerText))
                        headers[name] = headerText;
                }
            }

            var footerTypes = new[]
            {
                (HeaderFooterType.FooterPrimary, "primary"),
                (HeaderFooterType.FooterFirst, "firstPage"),
                (HeaderFooterType.FooterEven, "evenPage")
            };

            var footers = new Dictionary<string, string?>();
            foreach (var (type, name) in footerTypes)
            {
                var footer = section.HeadersFooters[type];
                if (footer != null)
                {
                    var footerText = footer.ToString(SaveFormat.Text).Trim();
                    if (!string.IsNullOrEmpty(footerText))
                        footers[name] = footerText;
                }
            }

            sectionsList.Add(new
            {
                sectionIndex = actualIndex,
                headers = headers.Count > 0 ? headers : null,
                footers = footers.Count > 0 ? footers : null
            });
        }

        var result = new
        {
            totalSections = doc.Sections.Count,
            queriedSectionIndex = sectionIndex,
            sections = sectionsList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Inserts text or field code. If text contains field codes like {PAGE}, {DATE}, {NUMPAGES}, etc.,
    ///     they will be inserted as actual fields instead of plain text.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="text">The text to insert, may contain field codes.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    private static void InsertTextOrField(DocumentBuilder builder, string text, string? fontName, string? fontNameAscii,
        string? fontNameFarEast, double? fontSize)
    {
        // Common field codes that should be converted to fields
        var fieldCodes = new HashSet<string>
        {
            "PAGE", "NUMPAGES", "DATE", "TIME", "AUTHOR", "FILENAME", "TITLE", "CREATEDATE", "SAVEDATE", "PRINTDATE"
        };

        // Pattern to match field codes like {PAGE}, {DATE}, etc.
        var fieldPattern = new Regex(@"\{([A-Z]+)\}", RegexOptions.IgnoreCase);
        var matches = fieldPattern.Matches(text);

        if (matches.Count == 0)
        {
            if (fontName != null || fontSize.HasValue)
            {
                // Apply font settings using FontHelper
                FontHelper.Word.ApplyFontSettings(
                    builder,
                    fontName,
                    fontNameAscii,
                    fontNameFarEast,
                    fontSize
                );
                builder.Write(text);
                // Restore to Normal style
                RestoreNormalFont(builder);
            }
            else
            {
                builder.Write(text);
            }

            return;
        }

        var lastIndex = 0;
        foreach (Match match in matches)
        {
            if (match.Index > lastIndex)
            {
                var textBefore = text.Substring(lastIndex, match.Index - lastIndex);
                if (!string.IsNullOrEmpty(textBefore))
                {
                    if (fontName != null || fontSize.HasValue)
                        // Apply font settings using FontHelper
                        FontHelper.Word.ApplyFontSettings(
                            builder,
                            fontName,
                            fontNameAscii,
                            fontNameFarEast,
                            fontSize
                        );

                    builder.Write(textBefore);
                    if (fontName != null || fontSize.HasValue)
                        RestoreNormalFont(builder);
                }
            }

            var fieldCode = match.Groups[1].Value.ToUpper();
            if (fieldCodes.Contains(fieldCode))
            {
                var fieldType = fieldCode switch
                {
                    "PAGE" => FieldType.FieldPage,
                    "NUMPAGES" => FieldType.FieldNumPages,
                    "DATE" => FieldType.FieldDate,
                    "TIME" => FieldType.FieldTime,
                    "AUTHOR" => FieldType.FieldAuthor,
                    "FILENAME" => FieldType.FieldFileName,
                    "TITLE" => FieldType.FieldTitle,
                    "CREATEDATE" => FieldType.FieldCreateDate,
                    "SAVEDATE" => FieldType.FieldSaveDate,
                    "PRINTDATE" => FieldType.FieldPrintDate,
                    _ => throw new ArgumentException($"Unknown field code: {fieldCode}")
                };

                try
                {
                    var field = builder.InsertField(fieldType, false);
                    field.Update();
                }
                catch (Exception ex)
                {
                    // Field insertion failed, restore font settings if they were modified
                    Console.Error.WriteLine(
                        $"[WARN] Failed to insert field '{fieldType}' in header/footer: {ex.Message}");
                    if (fontName != null || fontSize.HasValue)
                        // Apply font settings using FontHelper
                        FontHelper.Word.ApplyFontSettings(
                            builder,
                            fontName,
                            fontNameAscii,
                            fontNameFarEast,
                            fontSize
                        );

                    builder.Write(match.Value);
                    if (fontName != null || fontSize.HasValue)
                        RestoreNormalFont(builder);
                }
            }
            else
            {
                if (fontName != null || fontSize.HasValue)
                    // Apply font settings using FontHelper
                    FontHelper.Word.ApplyFontSettings(
                        builder,
                        fontName,
                        fontNameAscii,
                        fontNameFarEast,
                        fontSize
                    );

                builder.Write(match.Value);
                if (fontName != null || fontSize.HasValue)
                    RestoreNormalFont(builder);
            }

            lastIndex = match.Index + match.Length;
        }

        if (lastIndex < text.Length)
        {
            var textAfter = text.Substring(lastIndex);
            if (!string.IsNullOrEmpty(textAfter))
            {
                if (fontName != null || fontSize.HasValue)
                    // Apply font settings using FontHelper
                    FontHelper.Word.ApplyFontSettings(
                        builder,
                        fontName,
                        fontNameAscii,
                        fontNameFarEast,
                        fontSize
                    );

                builder.Write(textAfter);
                if (fontName != null || fontSize.HasValue)
                    RestoreNormalFont(builder);
            }
        }
    }

    /// <summary>
    ///     Restores DocumentBuilder font to Normal style.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    private static void RestoreNormalFont(DocumentBuilder builder)
    {
        builder.Font.Name = builder.Document.Styles[StyleIdentifier.Normal].Font.Name;
        builder.Font.Size = builder.Document.Styles[StyleIdentifier.Normal].Font.Size;
    }

    /// <summary>
    ///     Gets the appropriate HeaderFooterType based on the type string and whether it's a header or footer.
    /// </summary>
    /// <param name="typeStr">The type string (primary, firstpage, evenpages).</param>
    /// <param name="isHeader">Whether this is for a header (true) or footer (false).</param>
    /// <returns>The corresponding HeaderFooterType enum value.</returns>
    private static HeaderFooterType GetHeaderFooterType(string typeStr, bool isHeader)
    {
        return typeStr.ToLower() switch
        {
            "firstpage" => isHeader ? HeaderFooterType.HeaderFirst : HeaderFooterType.FooterFirst,
            "evenpages" => isHeader ? HeaderFooterType.HeaderEven : HeaderFooterType.FooterEven,
            _ => isHeader ? HeaderFooterType.HeaderPrimary : HeaderFooterType.FooterPrimary
        };
    }

    /// <summary>
    ///     Clears only text content from a header/footer, preserving shapes and images.
    /// </summary>
    /// <param name="headerFooter">The header or footer to clear text from.</param>
    private static void ClearTextOnly(HeaderFooter headerFooter)
    {
        // Remove paragraphs and runs but keep shapes
        List<Node> nodesToRemove = [];
        foreach (var para in headerFooter.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>())
        {
            // Check if paragraph contains only text (no shapes)
            var hasShapes = para.GetChildNodes(NodeType.Shape, true).Count > 0;
            if (!hasShapes)
            {
                nodesToRemove.Add(para);
            }
            else
            {
                // Remove only runs (text) but keep shapes
                var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
                foreach (var run in runs)
                    run.Remove();
            }
        }

        foreach (var node in nodesToRemove)
            node.Remove();
    }
}