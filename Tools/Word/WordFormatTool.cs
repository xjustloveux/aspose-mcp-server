using System.ComponentModel;
using System.Drawing;
using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for formatting text and paragraphs in Word documents
/// </summary>
[McpServerToolType]
public class WordFormatTool
{
    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordFormatTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordFormatTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_format")]
    [Description(
        @"Manage formatting in Word documents. Supports 6 operations: get_run_format, set_run_format, get_tab_stops, add_tab_stop, clear_tab_stops, set_paragraph_border.

Usage examples:
- Get run format: word_format(operation='get_run_format', path='doc.docx', paragraphIndex=0, runIndex=0)
- Get inherited format: word_format(operation='get_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, includeInherited=true)
- Set run format: word_format(operation='set_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true, fontSize=14)
- Reset color to auto: word_format(operation='set_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, color='auto')
- Get tab stops: word_format(operation='get_tab_stops', path='doc.docx', paragraphIndex=0)
- Add tab stop: word_format(operation='add_tab_stop', path='doc.docx', paragraphIndex=0, tabPosition=72, tabAlignment='center')
- Clear tab stops: word_format(operation='clear_tab_stops', path='doc.docx', paragraphIndex=0)
- Set paragraph border: word_format(operation='set_paragraph_border', path='doc.docx', paragraphIndex=0, borderPosition='all', lineStyle='single', lineWidth=1.0)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get_run_format': Get run formatting (required params: path, paragraphIndex, runIndex)
- 'set_run_format': Set run formatting (required params: path, paragraphIndex, runIndex)
- 'get_tab_stops': Get tab stops (required params: path, paragraphIndex)
- 'add_tab_stop': Add a tab stop (required params: path, paragraphIndex, tabPosition)
- 'clear_tab_stops': Clear tab stops (required params: path, paragraphIndex)
- 'set_paragraph_border': Set paragraph border (required params: path, paragraphIndex)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Paragraph index (0-based)")]
        int? paragraphIndex = null,
        [Description("Run index within paragraph (0-based, optional)")]
        int? runIndex = null,
        [Description("Section index (0-based, default: 0)")]
        int sectionIndex = 0,
        [Description(
            "Include inherited format from paragraph/style (for get_run_format, default: false). When true, shows the effective computed format.")]
        bool includeInherited = false,
        [Description("Font name (for set_run_format)")]
        string? fontName = null,
        [Description("Font name for ASCII characters (for set_run_format)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (for set_run_format)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (for set_run_format)")]
        double? fontSize = null,
        [Description("Bold text (for set_run_format)")]
        bool? bold = null,
        [Description("Italic text (for set_run_format)")]
        bool? italic = null,
        [Description("Underline text (for set_run_format)")]
        bool? underline = null,
        [Description("Font color hex (for set_run_format)")]
        string? color = null,
        [Description("Where to get tab stops from: header, footer, body (for get_tab_stops, default: body)")]
        string location = "body",
        [Description("Read tab stops from all paragraphs (for get_tab_stops, default: false)")]
        bool allParagraphs = false,
        [Description("Include tab stops from paragraph style (for get_tab_stops, default: true)")]
        bool includeStyle = true,
        [Description("Tab stop position in points (for add_tab_stop, required)")]
        double? tabPosition = null,
        [Description("Tab stop alignment (for add_tab_stop, default: left)")]
        string tabAlignment = "left",
        [Description("Tab stop leader character (for add_tab_stop, default: none)")]
        string tabLeader = "none",
        [Description(
            "Border position shortcut (for set_paragraph_border): 'all', 'top-bottom', 'left-right', 'box'. Overrides individual border flags.")]
        string? borderPosition = null,
        [Description("Show top border (for set_paragraph_border, default: false)")]
        bool borderTop = false,
        [Description("Show bottom border (for set_paragraph_border, default: false)")]
        bool borderBottom = false,
        [Description("Show left border (for set_paragraph_border, default: false)")]
        bool borderLeft = false,
        [Description("Show right border (for set_paragraph_border, default: false)")]
        bool borderRight = false,
        [Description("Border line style: none, single, double, dotted, dashed, thick (for set_paragraph_border)")]
        string lineStyle = "single",
        [Description("Border line width in points (for set_paragraph_border, default: 0.5)")]
        double lineWidth = 0.5,
        [Description("Border line color hex (for set_paragraph_border, default: 000000)")]
        string lineColor = "000000")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "get_run_format" => GetRunFormat(ctx, paragraphIndex ?? 0, runIndex, includeInherited),
            "set_run_format" => SetRunFormat(ctx, outputPath, paragraphIndex ?? 0, runIndex, fontName, fontNameAscii,
                fontNameFarEast, fontSize, bold, italic, underline, color),
            "get_tab_stops" => GetTabStops(ctx, location, paragraphIndex ?? 0, sectionIndex, allParagraphs,
                includeStyle),
            "add_tab_stop" => AddTabStop(ctx, outputPath, paragraphIndex ?? 0, tabPosition ?? 0, tabAlignment,
                tabLeader),
            "clear_tab_stops" => ClearTabStops(ctx, outputPath, paragraphIndex ?? 0),
            "set_paragraph_border" => SetParagraphBorder(ctx, outputPath, paragraphIndex ?? 0, borderPosition,
                borderTop, borderBottom, borderLeft, borderRight, lineStyle, lineWidth, lineColor),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets run format information.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="runIndex">The optional zero-based run index within the paragraph.</param>
    /// <param name="includeInherited">Whether to include inherited format from paragraph/style.</param>
    /// <returns>A JSON string containing the run format information.</returns>
    /// <exception cref="ArgumentException">Thrown when runIndex is out of range.</exception>
    private static string GetRunFormat(DocumentContext<Document> ctx, int paragraphIndex, int? runIndex,
        bool includeInherited)
    {
        var doc = ctx.Document;

        var para = GetTargetParagraph(doc, paragraphIndex);
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
                throw new ArgumentException(
                    $"runIndex {runIndex.Value} is out of range (paragraph #{paragraphIndex} has {runs.Count} Runs, valid range: 0-{runs.Count - 1})");

            var run = runs[runIndex.Value];
            var font = run.Font;
            var colorHex = $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}";
            var colorName = GetColorName(font.Color);

            object result;
            if (includeInherited)
                result = new
                {
                    paragraphIndex,
                    runIndex = runIndex.Value,
                    text = run.Text,
                    formatType = "inherited",
                    fontName = font.Name,
                    fontNameAscii = font.NameAscii,
                    fontNameFarEast = font.NameFarEast,
                    fontSize = font.Size,
                    bold = font.Bold,
                    italic = font.Italic,
                    underline = font.Underline.ToString(),
                    strikeThrough = font.StrikeThrough,
                    superscript = font.Superscript,
                    subscript = font.Subscript,
                    color = colorHex,
                    colorName,
                    isAutoColor = font.Color is { IsEmpty: true } or { R: 0, G: 0, B: 0, A: 0 }
                };
            else
                result = new
                {
                    paragraphIndex,
                    runIndex = runIndex.Value,
                    text = run.Text,
                    formatType = "explicit",
                    fontName = font.Name,
                    fontNameAscii = font.NameAscii,
                    fontNameFarEast = font.NameFarEast,
                    fontSize = font.Size,
                    bold = font.Bold,
                    italic = font.Italic,
                    underline = font.Underline.ToString(),
                    strikeThrough = font.StrikeThrough,
                    superscript = font.Superscript,
                    subscript = font.Subscript,
                    color = colorHex,
                    colorName,
                    isAutoColor = font.Color is { IsEmpty: true } or { R: 0, G: 0, B: 0, A: 0 }
                };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var runsList = runs.Select((run, i) =>
            {
                var font = run.Font;
                var colorHex = $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}";
                return new
                {
                    index = i,
                    text = run.Text,
                    fontNameAscii = font.NameAscii,
                    fontNameFarEast = font.NameFarEast,
                    fontSize = font.Size,
                    bold = font.Bold,
                    italic = font.Italic,
                    underline = font.Underline.ToString(),
                    strikeThrough = font.StrikeThrough,
                    superscript = font.Superscript,
                    subscript = font.Subscript,
                    color = colorHex,
                    colorName = GetColorName(font.Color)
                };
            }).ToList();

            var result = new
            {
                paragraphIndex,
                count = runs.Count,
                runs = runsList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }

    /// <summary>
    ///     Sets run format properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="runIndex">The optional zero-based run index within the paragraph.</param>
    /// <param name="fontName">The font name to apply.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether to apply bold formatting.</param>
    /// <param name="italic">Whether to apply italic formatting.</param>
    /// <param name="underline">Whether to apply underline formatting.</param>
    /// <param name="color">The font color in hex format or 'auto'.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when runIndex is out of range or paragraph has no runs.</exception>
    private static string SetRunFormat(DocumentContext<Document> ctx, string? outputPath, int paragraphIndex,
        int? runIndex,
        string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize,
        bool? bold, bool? italic, bool? underline, string? color)
    {
        var doc = ctx.Document;
        var para = GetTargetParagraph(doc, paragraphIndex);
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        // If paragraph has no runs and runIndex is specified, create a run
        if (runs.Count == 0 && runIndex.HasValue)
        {
            if (runIndex.Value != 0)
                throw new ArgumentException("Paragraph has no Run nodes, runIndex must be 0 to create a new Run");
            // Create a new run with empty text
            var newRun = new Run(doc);
            para.AppendChild(newRun);
            runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        }
        // If paragraph has no runs and no runIndex specified, create a run
        else if (runs.Count == 0)
        {
            var newRun = new Run(doc);
            para.AppendChild(newRun);
            runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        }

        List<Run> runsToFormat;
        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
                throw new ArgumentException(
                    $"runIndex must be between 0 and {runs.Count - 1} (paragraph has {runs.Count} Runs)");
            runsToFormat = [runs[runIndex.Value]];
        }
        else
        {
            runsToFormat = runs;
        }

        var isAutoColor = color?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;

        foreach (var run in runsToFormat)
        {
            // Apply font settings using FontHelper
            var underlineStr = underline.HasValue ? underline.Value ? "single" : "none" : null;

            if (isAutoColor)
                run.Font.Color = Color.Empty;
            else
                FontHelper.Word.ApplyFontSettings(
                    run,
                    fontName,
                    fontNameAscii,
                    fontNameFarEast,
                    fontSize,
                    bold,
                    italic,
                    underlineStr,
                    color
                );

            // Apply other font settings when auto color is set
            if (isAutoColor)
            {
                if (!string.IsNullOrEmpty(fontName)) run.Font.Name = fontName;
                if (!string.IsNullOrEmpty(fontNameAscii)) run.Font.NameAscii = fontNameAscii;
                if (!string.IsNullOrEmpty(fontNameFarEast)) run.Font.NameFarEast = fontNameFarEast;
                if (fontSize.HasValue) run.Font.Size = fontSize.Value;
                if (bold.HasValue) run.Font.Bold = bold.Value;
                if (italic.HasValue) run.Font.Italic = italic.Value;
                if (underline.HasValue) run.Font.Underline = underline.Value ? Underline.Single : Underline.None;
            }
        }

        ctx.Save(outputPath);
        var colorMsg = isAutoColor ? " (color reset to auto)" : "";
        var result = $"Run format updated{colorMsg}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets tab stops for a paragraph.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="location">The location to get tab stops from (header, footer, body).</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="sectionIndex">The zero-based section index.</param>
    /// <param name="allParagraphs">Whether to read tab stops from all paragraphs.</param>
    /// <param name="includeStyle">Whether to include tab stops from paragraph style.</param>
    /// <returns>A JSON string containing the tab stops information.</returns>
    /// <exception cref="ArgumentException">Thrown when section or paragraph index is out of range.</exception>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the specified header/footer is not found or no target
    ///     paragraphs exist.
    /// </exception>
    private static string GetTabStops(DocumentContext<Document> ctx, string location, int paragraphIndex,
        int sectionIndex, bool allParagraphs, bool includeStyle)
    {
        var doc = ctx.Document;

        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException(
                $"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");

        var section = doc.Sections[sectionIndex];

        List<Paragraph> targetParagraphs;
        string locationDesc;

        switch (location.ToLower())
        {
            case "header":
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header != null)
                {
                    var headerParas = header.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    targetParagraphs = allParagraphs ? headerParas :
                        headerParas.Count > 0 ? [headerParas[0]] : [];
                    locationDesc = "Header";
                }
                else
                {
                    throw new InvalidOperationException("Header not found");
                }

                break;

            case "footer":
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                {
                    var footerParas = footer.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    targetParagraphs = allParagraphs ? footerParas :
                        footerParas.Count > 0 ? [footerParas[0]] : [];
                    locationDesc = "Footer";
                }
                else
                {
                    throw new InvalidOperationException("Footer not found");
                }

                break;

            default:
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                if (allParagraphs)
                {
                    targetParagraphs = paragraphs;
                }
                else
                {
                    if (paragraphIndex >= paragraphs.Count)
                        throw new ArgumentException(
                            $"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count})");
                    targetParagraphs = [paragraphs[paragraphIndex]];
                }

                locationDesc = allParagraphs ? "Body" : $"Body Paragraph {paragraphIndex}";
                break;
        }

        if (targetParagraphs.Count == 0)
            throw new InvalidOperationException("No target paragraphs found");

        var allTabStops =
            new Dictionary<string, (double position, TabAlignment alignment, TabLeader leader, string source)>();

        for (var paraIdx = 0; paraIdx < targetParagraphs.Count; paraIdx++)
        {
            var para = targetParagraphs[paraIdx];
            var paraSource = allParagraphs ? $"Paragraph {paraIdx}" : "Paragraph";

            var paraTabStops = para.ParagraphFormat.TabStops;
            for (var i = 0; i < paraTabStops.Count; i++)
            {
                var tab = paraTabStops[i];
                var position = Math.Round(tab.Position, 2);
                var key = $"{position}_{tab.Alignment}";
                if (!allTabStops.ContainsKey(key))
                    allTabStops[key] = (position, tab.Alignment, tab.Leader, $"{paraSource} (Custom)");
            }

            if (includeStyle && para.ParagraphFormat.Style != null)
            {
                var currentStyle = para.ParagraphFormat.Style;
                List<Style> styleChain = [];

                while (currentStyle != null)
                {
                    styleChain.Add(currentStyle);
                    if (!string.IsNullOrEmpty(currentStyle.BaseStyleName))
                        try
                        {
                            var baseStyle = para.Document.Styles[currentStyle.BaseStyleName];
                            if (baseStyle != null && !styleChain.Contains(baseStyle))
                                currentStyle = baseStyle;
                            else
                                currentStyle = null;
                        }
                        catch (Exception ex)
                        {
                            currentStyle = null;
                            Console.Error.WriteLine($"[WARN] Error accessing paragraph style: {ex.Message}");
                        }
                    else
                        currentStyle = null;
                }

                foreach (var chainStyle in styleChain)
                    if (chainStyle.ParagraphFormat != null)
                    {
                        var styleTabStops = chainStyle.ParagraphFormat.TabStops;
                        for (var i = 0; i < styleTabStops.Count; i++)
                        {
                            var tab = styleTabStops[i];
                            var position = Math.Round(tab.Position, 2);
                            var key = $"{position}_{tab.Alignment}";

                            if (!allTabStops.ContainsKey(key))
                            {
                                var styleName = chainStyle == para.ParagraphFormat.Style
                                    ? chainStyle.Name
                                    : $"{para.ParagraphFormat.Style.Name} (Base: {chainStyle.Name})";
                                allTabStops[key] = (position, tab.Alignment, tab.Leader,
                                    $"{paraSource} (Style: {styleName})");
                            }
                        }
                    }
            }
        }

        // Build JSON response
        var tabStopsList = allTabStops
            .OrderBy(x => x.Value.position)
            .Select(kvp => new
            {
                positionPt = kvp.Value.position,
                positionCm = Math.Round(kvp.Value.position / 28.35, 2),
                alignment = kvp.Value.alignment.ToString(),
                leader = kvp.Value.leader.ToString(),
                kvp.Value.source
            })
            .ToList();

        var result = new
        {
            location,
            locationDescription = locationDesc,
            sectionIndex,
            paragraphIndex = location == "body" && !allParagraphs ? paragraphIndex : (int?)null,
            allParagraphs,
            includeStyle,
            paragraphCount = targetParagraphs.Count,
            count = allTabStops.Count,
            tabStops = tabStopsList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Sets paragraph border properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="borderPosition">The border position shortcut (all, top-bottom, left-right, box, none).</param>
    /// <param name="borderTop">Whether to show the top border.</param>
    /// <param name="borderBottom">Whether to show the bottom border.</param>
    /// <param name="borderLeft">Whether to show the left border.</param>
    /// <param name="borderRight">Whether to show the right border.</param>
    /// <param name="lineStyle">The border line style.</param>
    /// <param name="lineWidth">The border line width in points.</param>
    /// <param name="lineColor">The border line color in hex format.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when borderPosition is invalid.</exception>
    private static string SetParagraphBorder(DocumentContext<Document> ctx, string? outputPath, int paragraphIndex,
        string? borderPosition, bool borderTop, bool borderBottom, bool borderLeft, bool borderRight,
        string lineStyle, double lineWidth, string lineColor)
    {
        var doc = ctx.Document;
        var para = GetTargetParagraph(doc, paragraphIndex);
        var borders = para.ParagraphFormat.Borders;

        bool actualBorderTop, actualBorderBottom, actualBorderLeft, actualBorderRight;

        if (!string.IsNullOrEmpty(borderPosition))
        {
            // borderPosition overrides individual flags
            switch (borderPosition.ToLower())
            {
                case "all":
                case "box":
                    actualBorderTop = actualBorderBottom = actualBorderLeft = actualBorderRight = true;
                    break;
                case "top-bottom":
                    actualBorderTop = actualBorderBottom = true;
                    actualBorderLeft = actualBorderRight = false;
                    break;
                case "left-right":
                    actualBorderTop = actualBorderBottom = false;
                    actualBorderLeft = actualBorderRight = true;
                    break;
                case "none":
                    actualBorderTop = actualBorderBottom = actualBorderLeft = actualBorderRight = false;
                    break;
                default:
                    throw new ArgumentException(
                        $"Invalid borderPosition: {borderPosition}. Valid values: all, box, top-bottom, left-right, none");
            }
        }
        else
        {
            // Use individual flags
            actualBorderTop = borderTop;
            actualBorderBottom = borderBottom;
            actualBorderLeft = borderLeft;
            actualBorderRight = borderRight;
        }

        if (actualBorderTop)
        {
            borders.Top.LineStyle = GetLineStyle(lineStyle);
            borders.Top.LineWidth = lineWidth;
            borders.Top.Color = ColorHelper.ParseColor(lineColor);
        }
        else
        {
            borders.Top.LineStyle = LineStyle.None;
        }

        if (actualBorderBottom)
        {
            borders.Bottom.LineStyle = GetLineStyle(lineStyle);
            borders.Bottom.LineWidth = lineWidth;
            borders.Bottom.Color = ColorHelper.ParseColor(lineColor);
        }
        else
        {
            borders.Bottom.LineStyle = LineStyle.None;
        }

        if (actualBorderLeft)
        {
            borders.Left.LineStyle = GetLineStyle(lineStyle);
            borders.Left.LineWidth = lineWidth;
            borders.Left.Color = ColorHelper.ParseColor(lineColor);
        }
        else
        {
            borders.Left.LineStyle = LineStyle.None;
        }

        if (actualBorderRight)
        {
            borders.Right.LineStyle = GetLineStyle(lineStyle);
            borders.Right.LineWidth = lineWidth;
            borders.Right.Color = ColorHelper.ParseColor(lineColor);
        }
        else
        {
            borders.Right.LineStyle = LineStyle.None;
        }

        ctx.Save(outputPath);

        List<string> enabledBorders = [];
        if (actualBorderTop) enabledBorders.Add("Top");
        if (actualBorderBottom) enabledBorders.Add("Bottom");
        if (actualBorderLeft) enabledBorders.Add("Left");
        if (actualBorderRight) enabledBorders.Add("Right");

        var bordersDesc = enabledBorders.Count > 0 ? string.Join(", ", enabledBorders) : "None";

        var result = $"Successfully set paragraph {paragraphIndex} borders: {bordersDesc}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Adds a tab stop to a paragraph.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="tabPosition">The tab stop position in points.</param>
    /// <param name="tabAlignmentStr">The tab stop alignment (left, center, right, decimal, bar).</param>
    /// <param name="tabLeaderStr">The tab stop leader character (none, dots, dashes, line, heavy, middledot).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string AddTabStop(DocumentContext<Document> ctx, string? outputPath, int paragraphIndex,
        double tabPosition, string tabAlignmentStr, string tabLeaderStr)
    {
        var doc = ctx.Document;
        var para = GetTargetParagraph(doc, paragraphIndex);

        var tabAlignment = tabAlignmentStr.ToLower() switch
        {
            "center" => TabAlignment.Center,
            "right" => TabAlignment.Right,
            "decimal" => TabAlignment.Decimal,
            "bar" => TabAlignment.Bar,
            _ => TabAlignment.Left
        };

        var tabLeader = tabLeaderStr.ToLower() switch
        {
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            "heavy" => TabLeader.Heavy,
            "middledot" => TabLeader.MiddleDot,
            _ => TabLeader.None
        };

        para.ParagraphFormat.TabStops.Add(new TabStop(tabPosition, tabAlignment, tabLeader));

        ctx.Save(outputPath);
        var result = $"Tab stop added at {tabPosition}pt ({tabAlignmentStr}, {tabLeaderStr})\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Clears tab stops from a paragraph.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string ClearTabStops(DocumentContext<Document> ctx, string? outputPath, int paragraphIndex)
    {
        var doc = ctx.Document;
        var para = GetTargetParagraph(doc, paragraphIndex);

        var count = para.ParagraphFormat.TabStops.Count;
        para.ParagraphFormat.TabStops.Clear();

        ctx.Save(outputPath);
        var result = $"Cleared {count} tab stop(s) from paragraph {paragraphIndex}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Converts line style string to LineStyle enum.
    /// </summary>
    /// <param name="style">The line style string (none, single, double, dotted, dashed, thick).</param>
    /// <returns>The corresponding LineStyle enum value.</returns>
    private static LineStyle GetLineStyle(string style)
    {
        return style.ToLower() switch
        {
            "none" => LineStyle.None,
            "single" => LineStyle.Single,
            "double" => LineStyle.Double,
            "dotted" => LineStyle.Dot,
            "dashed" => LineStyle.Single,
            "thick" => LineStyle.Thick,
            _ => LineStyle.Single
        };
    }

    /// <summary>
    ///     Gets target paragraph using flat list.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index, or -1 for the last paragraph.</param>
    /// <returns>The target paragraph.</returns>
    /// <exception cref="ArgumentException">Thrown when the document has no paragraphs or the index is out of range.</exception>
    private static Paragraph GetTargetParagraph(Document doc, int paragraphIndex)
    {
        var allParas = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (allParas.Count == 0)
            throw new ArgumentException("Document has no paragraphs.");

        if (paragraphIndex == -1)
            return allParas[^1]; // Last paragraph

        if (paragraphIndex < 0 || paragraphIndex >= allParas.Count)
            throw new ArgumentException(
                $"paragraphIndex must be between 0 and {allParas.Count - 1}, or use -1 for last paragraph");

        return allParas[paragraphIndex];
    }

    /// <summary>
    ///     Gets human-readable color name.
    /// </summary>
    /// <param name="color">The color to get the name for.</param>
    /// <returns>The human-readable color name.</returns>
    private static string GetColorName(Color color)
    {
        if (color is { IsEmpty: true } or { R: 0, G: 0, B: 0, A: 0 })
            return "Auto/Black";

        // Check for known colors using pattern matching
        if (color is { R: 255, G: 0, B: 0 }) return "Red";
        if (color is { R: 0, G: 255, B: 0 }) return "Green";
        if (color is { R: 0, G: 0, B: 255 }) return "Blue";
        if (color is { R: 255, G: 255, B: 0 }) return "Yellow";
        if (color is { R: 255, G: 0, B: 255 }) return "Magenta";
        if (color is { R: 0, G: 255, B: 255 }) return "Cyan";
        if (color is { R: 255, G: 255, B: 255 }) return "White";
        if (color is { R: 128, G: 128, B: 128 }) return "Gray";
        if (color is { R: 255, G: 165, B: 0 }) return "Orange";
        if (color is { R: 128, G: 0, B: 128 }) return "Purple";

        // Try to get the named color
        if (color.IsKnownColor)
            return color.Name;

        return "Custom";
    }
}