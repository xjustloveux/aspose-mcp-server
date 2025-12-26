using System.Drawing;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for formatting text and paragraphs in Word documents
/// </summary>
public class WordFormatTool : IAsposeTool
{
    public string Description =>
        @"Manage formatting in Word documents. Supports 6 operations: get_run_format, set_run_format, get_tab_stops, add_tab_stop, clear_tab_stops, set_paragraph_border.

Usage examples:
- Get run format: word_format(operation='get_run_format', path='doc.docx', paragraphIndex=0, runIndex=0)
- Get inherited format: word_format(operation='get_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, includeInherited=true)
- Set run format: word_format(operation='set_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true, fontSize=14)
- Reset color to auto: word_format(operation='set_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, color='auto')
- Get tab stops: word_format(operation='get_tab_stops', path='doc.docx', paragraphIndex=0)
- Add tab stop: word_format(operation='add_tab_stop', path='doc.docx', paragraphIndex=0, tabPosition=72, tabAlignment='center')
- Clear tab stops: word_format(operation='clear_tab_stops', path='doc.docx', paragraphIndex=0)
- Set paragraph border: word_format(operation='set_paragraph_border', path='doc.docx', paragraphIndex=0, borderPosition='all', lineStyle='single', lineWidth=1.0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get_run_format': Get run formatting (required params: path, paragraphIndex, runIndex)
- 'set_run_format': Set run formatting (required params: path, paragraphIndex, runIndex)
- 'get_tab_stops': Get tab stops (required params: path, paragraphIndex)
- 'add_tab_stop': Add a tab stop (required params: path, paragraphIndex, tabPosition)
- 'clear_tab_stops': Clear tab stops (required params: path, paragraphIndex)
- 'set_paragraph_border': Set paragraph border (required params: path, paragraphIndex)",
                @enum = new[]
                {
                    "get_run_format", "set_run_format", "get_tab_stops", "add_tab_stop", "clear_tab_stops",
                    "set_paragraph_border"
                }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based)"
            },
            runIndex = new
            {
                type = "number",
                description = "Run index within paragraph (0-based, optional)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, default: 0)"
            },
            includeInherited = new
            {
                type = "boolean",
                description =
                    "Include inherited format from paragraph/style (for get_run_format, default: false). When true, shows the effective computed format."
            },
            fontName = new
            {
                type = "string",
                description = "Font name (for set_run_format)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (for set_run_format)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (for set_run_format)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (for set_run_format)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (for set_run_format)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (for set_run_format)"
            },
            underline = new
            {
                type = "boolean",
                description = "Underline text (for set_run_format)"
            },
            color = new
            {
                type = "string",
                description = "Font color hex (for set_run_format)"
            },
            location = new
            {
                type = "string",
                description = "Where to get tab stops from: header, footer, body (for get_tab_stops, default: body)",
                @enum = new[] { "header", "footer", "body" }
            },
            allParagraphs = new
            {
                type = "boolean",
                description = "Read tab stops from all paragraphs (for get_tab_stops, default: false)"
            },
            includeStyle = new
            {
                type = "boolean",
                description = "Include tab stops from paragraph style (for get_tab_stops, default: true)"
            },
            tabPosition = new
            {
                type = "number",
                description = "Tab stop position in points (for add_tab_stop, required)"
            },
            tabAlignment = new
            {
                type = "string",
                description = "Tab stop alignment (for add_tab_stop, default: left)",
                @enum = new[] { "left", "center", "right", "decimal", "bar" }
            },
            tabLeader = new
            {
                type = "string",
                description = "Tab stop leader character (for add_tab_stop, default: none)",
                @enum = new[] { "none", "dots", "dashes", "line", "heavy", "middleDot" }
            },
            borderPosition = new
            {
                type = "string",
                description =
                    "Border position shortcut (for set_paragraph_border): 'all', 'top-bottom', 'left-right', 'box'. Overrides individual border flags.",
                @enum = new[] { "all", "top-bottom", "left-right", "box", "none" }
            },
            borderTop = new
            {
                type = "boolean",
                description = "Show top border (for set_paragraph_border, default: false)"
            },
            borderBottom = new
            {
                type = "boolean",
                description = "Show bottom border (for set_paragraph_border, default: false)"
            },
            borderLeft = new
            {
                type = "boolean",
                description = "Show left border (for set_paragraph_border, default: false)"
            },
            borderRight = new
            {
                type = "boolean",
                description = "Show right border (for set_paragraph_border, default: false)"
            },
            lineStyle = new
            {
                type = "string",
                description =
                    "Border line style: none, single, double, dotted, dashed, thick (for set_paragraph_border)",
                @enum = new[] { "none", "single", "double", "dotted", "dashed", "thick" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Border line width in points (for set_paragraph_border, default: 0.5)"
            },
            lineColor = new
            {
                type = "string",
                description = "Border line color hex (for set_paragraph_border, default: 000000)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "get_run_format" => await GetRunFormat(path, arguments),
            "set_run_format" => await SetRunFormat(path, outputPath, arguments),
            "get_tab_stops" => await GetTabStops(path, arguments),
            "add_tab_stop" => await AddTabStop(path, outputPath, arguments),
            "clear_tab_stops" => await ClearTabStops(path, outputPath, arguments),
            "set_paragraph_border" => await SetParagraphBorder(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets run format information
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex, runIndex, optional sectionIndex, includeInherited</param>
    /// <returns>Formatted string with run format details</returns>
    private Task<string> GetRunFormat(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var runIndex = ArgumentHelper.GetIntNullable(arguments, "runIndex");
            var includeInherited = ArgumentHelper.GetBool(arguments, "includeInherited", false);

            var doc = new Document(path);

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
        });
    }

    /// <summary>
    ///     Sets run format properties
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex, runIndex, optional formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> SetRunFormat(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var runIndex = ArgumentHelper.GetIntNullable(arguments, "runIndex");
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
            var underline = ArgumentHelper.GetBoolNullable(arguments, "underline");
            var color = ArgumentHelper.GetStringNullable(arguments, "color");

            var doc = new Document(path);
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

            doc.Save(outputPath);
            var colorMsg = isAutoColor ? " (color reset to auto)" : "";
            return $"Run format updated{colorMsg}: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets tab stops for a paragraph
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional sectionIndex, location</param>
    /// <returns>Formatted string with tab stops</returns>
    private Task<string> GetTabStops(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var location = ArgumentHelper.GetString(arguments, "location", "body");
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", 0);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var allParagraphs = ArgumentHelper.GetBool(arguments, "allParagraphs", false);
            var includeStyle = ArgumentHelper.GetBool(arguments, "includeStyle", true);

            var doc = new Document(path);

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
                            headerParas.Count > 0 ? new List<Paragraph> { headerParas[0] } : new List<Paragraph>();
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
                            footerParas.Count > 0 ? new List<Paragraph> { footerParas[0] } : new List<Paragraph>();
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
                    var styleChain = new List<Style>();

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
        });
    }

    /// <summary>
    ///     Sets paragraph border properties
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex, optional border properties, borderPosition</param>
    /// <returns>Success message</returns>
    private Task<string> SetParagraphBorder(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");

            var doc = new Document(path);
            var para = GetTargetParagraph(doc, paragraphIndex);
            var borders = para.ParagraphFormat.Borders;

            var defaultLineStyle = ArgumentHelper.GetString(arguments, "lineStyle", "single");
            var defaultLineWidth = ArgumentHelper.GetDouble(arguments, "lineWidth", "lineWidth", 0.5);
            var defaultLineColor = ArgumentHelper.GetString(arguments, "lineColor", "000000");
            var borderPosition = ArgumentHelper.GetStringNullable(arguments, "borderPosition");
            bool borderTop, borderBottom, borderLeft, borderRight;

            if (!string.IsNullOrEmpty(borderPosition))
            {
                // borderPosition overrides individual flags
                switch (borderPosition.ToLower())
                {
                    case "all":
                    case "box":
                        borderTop = borderBottom = borderLeft = borderRight = true;
                        break;
                    case "top-bottom":
                        borderTop = borderBottom = true;
                        borderLeft = borderRight = false;
                        break;
                    case "left-right":
                        borderTop = borderBottom = false;
                        borderLeft = borderRight = true;
                        break;
                    case "none":
                        borderTop = borderBottom = borderLeft = borderRight = false;
                        break;
                    default:
                        throw new ArgumentException(
                            $"Invalid borderPosition: {borderPosition}. Valid values: all, box, top-bottom, left-right, none");
                }
            }
            else
            {
                // Use individual flags
                borderTop = ArgumentHelper.GetBool(arguments, "borderTop", false);
                borderBottom = ArgumentHelper.GetBool(arguments, "borderBottom", false);
                borderLeft = ArgumentHelper.GetBool(arguments, "borderLeft", false);
                borderRight = ArgumentHelper.GetBool(arguments, "borderRight", false);
            }

            if (borderTop)
            {
                borders.Top.LineStyle = GetLineStyle(defaultLineStyle);
                borders.Top.LineWidth = defaultLineWidth;
                borders.Top.Color = ColorHelper.ParseColor(defaultLineColor);
            }
            else
            {
                borders.Top.LineStyle = LineStyle.None;
            }

            if (borderBottom)
            {
                borders.Bottom.LineStyle = GetLineStyle(defaultLineStyle);
                borders.Bottom.LineWidth = defaultLineWidth;
                borders.Bottom.Color = ColorHelper.ParseColor(defaultLineColor);
            }
            else
            {
                borders.Bottom.LineStyle = LineStyle.None;
            }

            if (borderLeft)
            {
                borders.Left.LineStyle = GetLineStyle(defaultLineStyle);
                borders.Left.LineWidth = defaultLineWidth;
                borders.Left.Color = ColorHelper.ParseColor(defaultLineColor);
            }
            else
            {
                borders.Left.LineStyle = LineStyle.None;
            }

            if (borderRight)
            {
                borders.Right.LineStyle = GetLineStyle(defaultLineStyle);
                borders.Right.LineWidth = defaultLineWidth;
                borders.Right.Color = ColorHelper.ParseColor(defaultLineColor);
            }
            else
            {
                borders.Right.LineStyle = LineStyle.None;
            }

            doc.Save(outputPath);

            var enabledBorders = new List<string>();
            if (borderTop) enabledBorders.Add("Top");
            if (borderBottom) enabledBorders.Add("Bottom");
            if (borderLeft) enabledBorders.Add("Left");
            if (borderRight) enabledBorders.Add("Right");

            var bordersDesc = enabledBorders.Count > 0 ? string.Join(", ", enabledBorders) : "None";

            return $"Successfully set paragraph {paragraphIndex} borders: {bordersDesc}";
        });
    }

    /// <summary>
    ///     Adds a tab stop to a paragraph
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex, tabPosition, optional tabAlignment, tabLeader</param>
    /// <returns>Success message</returns>
    private Task<string> AddTabStop(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var tabPosition = ArgumentHelper.GetDouble(arguments, "tabPosition", "tabPosition");
            var tabAlignmentStr = ArgumentHelper.GetString(arguments, "tabAlignment", "left");
            var tabLeaderStr = ArgumentHelper.GetString(arguments, "tabLeader", "none");

            var doc = new Document(path);
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

            doc.Save(outputPath);
            return $"Tab stop added at {tabPosition}pt ({tabAlignmentStr}, {tabLeaderStr}): {outputPath}";
        });
    }

    /// <summary>
    ///     Clears tab stops from a paragraph
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing paragraphIndex</param>
    /// <returns>Success message with count of cleared tab stops</returns>
    private Task<string> ClearTabStops(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");

            var doc = new Document(path);
            var para = GetTargetParagraph(doc, paragraphIndex);

            var count = para.ParagraphFormat.TabStops.Count;
            para.ParagraphFormat.TabStops.Clear();

            doc.Save(outputPath);
            return $"Cleared {count} tab stop(s) from paragraph {paragraphIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Converts line style string to LineStyle enum
    /// </summary>
    /// <param name="style">Line style string (none, single, double, dotted, dashed, thick)</param>
    /// <returns>LineStyle enum value</returns>
    private LineStyle GetLineStyle(string style)
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
    ///     Gets target paragraph using flat list
    /// </summary>
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
    ///     Gets human-readable color name
    /// </summary>
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