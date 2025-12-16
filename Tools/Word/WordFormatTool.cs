using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using System.Linq;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordFormatTool : IAsposeTool
{
    public string Description => @"Manage formatting in Word documents. Supports 4 operations: get_run_format, set_run_format, get_tab_stops, set_paragraph_border.

Usage examples:
- Get run format: word_format(operation='get_run_format', path='doc.docx', paragraphIndex=0, runIndex=0)
- Set run format: word_format(operation='set_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true, fontSize=14)
- Get tab stops: word_format(operation='get_tab_stops', path='doc.docx', paragraphIndex=0)
- Set paragraph border: word_format(operation='set_paragraph_border', path='doc.docx', paragraphIndex=0, borderType='all', style='single', width=1.0)";

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
- 'set_paragraph_border': Set paragraph border (required params: path, paragraphIndex)",
                @enum = new[] { "get_run_format", "set_run_format", "get_tab_stops", "set_paragraph_border" }
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
                description = "Border line style: none, single, double, dotted, dashed, thick (for set_paragraph_border)",
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "get_run_format" => await GetRunFormat(arguments),
            "set_run_format" => await SetRunFormat(arguments),
            "get_tab_stops" => await GetTabStops(arguments),
            "set_paragraph_border" => await SetParagraphBorder(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Gets run format information
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, paragraphIndex, runIndex, optional sectionIndex</param>
    /// <returns>Formatted string with run format details</returns>
    private async Task<string> GetRunFormat(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
        var runIndex = ArgumentHelper.GetIntNullable(arguments, "runIndex");
        var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

        var doc = new Document(path);
        
        // Handle paragraphIndex=-1 (document end)
        if (paragraphIndex == -1)
        {
            var lastSection = doc.LastSection;
            var bodyParagraphs = lastSection.Body.GetChildNodes(NodeType.Paragraph, false);
            if (bodyParagraphs.Count > 0)
            {
                paragraphIndex = bodyParagraphs.Count - 1;
                sectionIndex = doc.Sections.Count - 1;
            }
            else
            {
                throw new ArgumentException("Cannot get run format: document has no paragraphs.");
            }
        }
        
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[sectionIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

        var para = paragraphs[paragraphIndex];
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var sb = new StringBuilder();

        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
                throw new ArgumentException($"runIndex {runIndex.Value} is out of range (paragraph #{paragraphIndex} has {runs.Count} Runs, valid range: 0-{runs.Count - 1})");

            var run = runs[runIndex.Value];
            sb.AppendLine($"=== Run {runIndex.Value} Format ===");
            sb.AppendLine($"  Text: {run.Text}");
            sb.AppendLine($"  Font Name: {run.Font.Name}");
            sb.AppendLine($"  Font Name (ASCII): {run.Font.NameAscii}");
            sb.AppendLine($"  Font Name (Far East): {run.Font.NameFarEast}");
            sb.AppendLine($"  Font Size: {run.Font.Size} pt");
            sb.AppendLine($"  Bold: {run.Font.Bold}");
            sb.AppendLine($"  Italic: {run.Font.Italic}");
            sb.AppendLine($"  Underline: {run.Font.Underline}");
            sb.AppendLine($"  StrikeThrough: {run.Font.StrikeThrough}");
            sb.AppendLine($"  Superscript: {run.Font.Superscript}");
            sb.AppendLine($"  Subscript: {run.Font.Subscript}");
            sb.AppendLine($"  Color: #{run.Font.Color.R:X2}{run.Font.Color.G:X2}{run.Font.Color.B:X2}");
        }
        else
        {
            sb.AppendLine($"=== Runs in Paragraph {paragraphIndex} ({runs.Count}) ===");
            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                sb.AppendLine($"\n[{i}] Text: {run.Text}");
                sb.AppendLine($"    Font: {run.Font.NameAscii}/{run.Font.NameFarEast}, Size: {run.Font.Size}pt");
                sb.AppendLine($"    Bold: {run.Font.Bold}, Italic: {run.Font.Italic}");
                sb.AppendLine($"    Underline: {run.Font.Underline}");
                if (run.Font.StrikeThrough) sb.AppendLine($"    StrikeThrough: True");
                if (run.Font.Superscript) sb.AppendLine($"    Superscript: True");
                if (run.Font.Subscript) sb.AppendLine($"    Subscript: True");
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    /// Sets run format properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, paragraphIndex, runIndex, optional formatting options, sectionIndex, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> SetRunFormat(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
        var runIndex = ArgumentHelper.GetIntNullable(arguments, "runIndex");
        var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
        var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
        var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
        var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
        var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
        var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
        var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
        var underline = ArgumentHelper.GetBoolNullable(arguments, "underline");
        var color = ArgumentHelper.GetStringNullable(arguments, "color");

        var doc = new Document(path);
        
        // Handle paragraphIndex=-1 (document end)
        if (paragraphIndex == -1)
        {
            var lastSection = doc.LastSection;
            var bodyParagraphs = lastSection.Body.GetChildNodes(NodeType.Paragraph, false);
            if (bodyParagraphs.Count > 0)
            {
                paragraphIndex = bodyParagraphs.Count - 1;
                sectionIndex = doc.Sections.Count - 1;
            }
            else
            {
                throw new ArgumentException("Cannot get run format: document has no paragraphs.");
            }
        }
        
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[sectionIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

        var para = paragraphs[paragraphIndex];
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        // If paragraph has no runs and runIndex is specified, create a run
        if (runs.Count == 0 && runIndex.HasValue)
        {
            if (runIndex.Value != 0)
            {
                throw new ArgumentException($"Paragraph has no Run nodes, runIndex must be 0 to create a new Run");
            }
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
                throw new ArgumentException($"runIndex must be between 0 and {runs.Count - 1} (paragraph has {runs.Count} Runs)");
            runsToFormat = new List<Run> { runs[runIndex.Value] };
        }
        else
        {
            runsToFormat = runs;
        }

        foreach (var run in runsToFormat)
        {
            if (!string.IsNullOrEmpty(fontName)) run.Font.Name = fontName;
            if (!string.IsNullOrEmpty(fontNameAscii)) run.Font.NameAscii = fontNameAscii;
            if (!string.IsNullOrEmpty(fontNameFarEast)) run.Font.NameFarEast = fontNameFarEast;
            if (fontSize.HasValue) run.Font.Size = fontSize.Value;
            if (bold.HasValue) run.Font.Bold = bold.Value;
            if (italic.HasValue) run.Font.Italic = italic.Value;
            if (underline.HasValue) run.Font.Underline = underline.Value ? Underline.Single : Underline.None;
            if (!string.IsNullOrEmpty(color))
            {
                try
                {
                    run.Font.Color = ColorHelper.ParseColor(color);
                }
                catch
                {
                    // Ignore invalid color
                }
            }
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Run format updated: {outputPath}");
    }

    /// <summary>
    /// Gets tab stops for a paragraph
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, paragraphIndex, optional sectionIndex</param>
    /// <returns>Formatted string with tab stops</returns>
    private async Task<string> GetTabStops(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var location = ArgumentHelper.GetString(arguments, "location", "body");
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", 0);
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
        var allParagraphs = ArgumentHelper.GetBool(arguments, "allParagraphs", false);
        var includeStyle = ArgumentHelper.GetBool(arguments, "includeStyle");

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
        
        var section = doc.Sections[sectionIndex];
        var result = new StringBuilder();
        
        result.AppendLine($"=== Tab Stops Information ===");
        result.AppendLine($"Location: {location}");
        if (location == "body")
            result.AppendLine($"Paragraph Index: {paragraphIndex}");
        result.AppendLine($"Section Index: {sectionIndex}");
        result.AppendLine();

        List<Paragraph> targetParagraphs = new List<Paragraph>();
        string locationDesc = "";

        switch (location.ToLower())
        {
            case "header":
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header != null)
                {
                    var headerParas = header.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    targetParagraphs = allParagraphs ? headerParas : (headerParas.Count > 0 ? new List<Paragraph> { headerParas[0] } : new List<Paragraph>());
                    locationDesc = "Header";
                }
                else
                    throw new InvalidOperationException("Header not found");
                break;

            case "footer":
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                {
                    var footerParas = footer.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    targetParagraphs = allParagraphs ? footerParas : (footerParas.Count > 0 ? new List<Paragraph> { footerParas[0] } : new List<Paragraph>());
                    locationDesc = "Footer";
                }
                else
                    throw new InvalidOperationException("Footer not found");
                break;

            case "body":
            default:
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                if (allParagraphs)
                    targetParagraphs = paragraphs;
                else
                {
                    if (paragraphIndex >= paragraphs.Count)
                        throw new ArgumentException($"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count})");
                    targetParagraphs = new List<Paragraph> { paragraphs[paragraphIndex] };
                }
                locationDesc = allParagraphs ? "Body" : $"Body Paragraph {paragraphIndex}";
                break;
        }

        if (targetParagraphs.Count == 0)
            throw new InvalidOperationException("No target paragraphs found");

        var allTabStops = new Dictionary<string, (double position, TabAlignment alignment, TabLeader leader, string source)>();
        
        for (int paraIdx = 0; paraIdx < targetParagraphs.Count; paraIdx++)
        {
            var para = targetParagraphs[paraIdx];
            var paraSource = allParagraphs ? $"Paragraph {paraIdx}" : "Paragraph";
            
            var paraTabStops = para.ParagraphFormat.TabStops;
            for (int i = 0; i < paraTabStops.Count; i++)
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
                    {
                        try
                        {
                            var baseStyle = para.Document.Styles[currentStyle.BaseStyleName];
                            if (baseStyle != null && !styleChain.Contains(baseStyle))
                                currentStyle = baseStyle;
                            else
                                currentStyle = null;
                        }
                        catch
                        {
                            currentStyle = null;
                        }
                    }
                    else
                        currentStyle = null;
                }
                
                foreach (var chainStyle in styleChain)
                {
                    if (chainStyle.ParagraphFormat != null)
                    {
                        var styleTabStops = chainStyle.ParagraphFormat.TabStops;
                        for (int i = 0; i < styleTabStops.Count; i++)
                        {
                            var tab = styleTabStops[i];
                            var position = Math.Round(tab.Position, 2);
                            var key = $"{position}_{tab.Alignment}";
                            
                            if (!allTabStops.ContainsKey(key))
                            {
                                var styleName = chainStyle == para.ParagraphFormat.Style 
                                    ? chainStyle.Name 
                                    : $"{para.ParagraphFormat.Style.Name} (Base: {chainStyle.Name})";
                                allTabStops[key] = (position, tab.Alignment, tab.Leader, $"{paraSource} (Style: {styleName})");
                            }
                        }
                    }
                }
            }
        }
        
        result.AppendLine($"【Tab Stops in {locationDesc}】");
        if (allParagraphs)
            result.AppendLine($"Read Range: All paragraphs ({targetParagraphs.Count})");
        if (includeStyle)
            result.AppendLine($"Include Style Tab Stops: Yes");
        result.AppendLine();
        
        if (allTabStops.Count == 0)
        {
            result.AppendLine("  No Tab Stops");
        }
        else
        {
            result.AppendLine($"  Total {allTabStops.Count} Tab Stop(s):");
            result.AppendLine();
            
            int idx = 1;
            foreach (var kvp in allTabStops.OrderBy(x => x.Value.position))
            {
                var (position, alignment, leader, source) = kvp.Value;
                result.AppendLine($"  Tab Stop {idx}:");
                result.AppendLine($"    Position: {position:F2} pt ({position / 28.35:F2} cm)");
                result.AppendLine($"    Alignment: {alignment}");
                result.AppendLine($"    Leader: {leader}");
                result.AppendLine($"    Source: {source}");
                result.AppendLine();
                idx++;
            }
        }

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    /// Sets paragraph border properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, paragraphIndex, optional border properties, sectionIndex, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> SetParagraphBorder(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
        
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        
        if (paragraphIndex >= paragraphs.Count)
            throw new ArgumentException($"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count})");
        
        var para = paragraphs[paragraphIndex];
        var borders = para.ParagraphFormat.Borders;
        
        var defaultLineStyle = ArgumentHelper.GetString(arguments, "lineStyle", "single");
        var defaultLineWidth = ArgumentHelper.GetDouble(arguments, "lineWidth", "lineWidth", 0.5);
        var defaultLineColor = ArgumentHelper.GetString(arguments, "lineColor", "000000");
        
        if (ArgumentHelper.GetBool(arguments, "borderTop", false))
        {
            borders.Top.LineStyle = GetLineStyle(defaultLineStyle);
            borders.Top.LineWidth = defaultLineWidth;
            borders.Top.Color = ColorHelper.ParseColor(defaultLineColor);
        }
        else
            borders.Top.LineStyle = LineStyle.None;
        
        if (ArgumentHelper.GetBool(arguments, "borderBottom", false))
        {
            borders.Bottom.LineStyle = GetLineStyle(defaultLineStyle);
            borders.Bottom.LineWidth = defaultLineWidth;
            borders.Bottom.Color = ColorHelper.ParseColor(defaultLineColor);
        }
        else
            borders.Bottom.LineStyle = LineStyle.None;
        
        if (ArgumentHelper.GetBool(arguments, "borderLeft", false))
        {
            borders.Left.LineStyle = GetLineStyle(defaultLineStyle);
            borders.Left.LineWidth = defaultLineWidth;
            borders.Left.Color = ColorHelper.ParseColor(defaultLineColor);
        }
        else
            borders.Left.LineStyle = LineStyle.None;
        
        if (ArgumentHelper.GetBool(arguments, "borderRight", false))
        {
            borders.Right.LineStyle = GetLineStyle(defaultLineStyle);
            borders.Right.LineWidth = defaultLineWidth;
            borders.Right.Color = ColorHelper.ParseColor(defaultLineColor);
        }
        else
            borders.Right.LineStyle = LineStyle.None;
        
        doc.Save(outputPath);
        
        var enabledBorders = new List<string>();
        if (ArgumentHelper.GetBool(arguments, "borderTop", false)) enabledBorders.Add("Top");
        if (ArgumentHelper.GetBool(arguments, "borderBottom", false)) enabledBorders.Add("Bottom");
        if (ArgumentHelper.GetBool(arguments, "borderLeft", false)) enabledBorders.Add("Left");
        if (ArgumentHelper.GetBool(arguments, "borderRight", false)) enabledBorders.Add("Right");
        
        var bordersDesc = enabledBorders.Count > 0 ? string.Join(", ", enabledBorders) : "None";
        
        return await Task.FromResult($"Successfully set paragraph {paragraphIndex} borders: {bordersDesc}");
    }

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

}

