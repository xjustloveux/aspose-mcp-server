using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.IO;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordShapeTool : IAsposeTool
{
    public string Description => "Manage shapes in Word documents (add line, add/edit/get textbox, set textbox border, add chart)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border, add_chart",
                @enum = new[] { "add_line", "add_textbox", "get_textboxes", "edit_textbox_content", "set_textbox_border", "add_chart" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            location = new
            {
                type = "string",
                description = "Where to add line: body, header, footer (for add_line, default: body)",
                @enum = new[] { "body", "header", "footer" }
            },
            position = new
            {
                type = "string",
                description = "Position: start, end (for add_line, default: end)",
                @enum = new[] { "start", "end" }
            },
            lineStyle = new
            {
                type = "string",
                description = "Line style: border, shape (for add_line, default: shape)",
                @enum = new[] { "border", "shape" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Line width in points (for add_line, default: 1.0)"
            },
            lineColor = new
            {
                type = "string",
                description = "Line color hex (for add_line, default: 000000)"
            },
            width = new
            {
                type = "number",
                description = "Line length in points (for add_line, optional)"
            },
            text = new
            {
                type = "string",
                description = "Text content (for add_textbox, edit_textbox_content)"
            },
            textboxWidth = new
            {
                type = "number",
                description = "Textbox width in points (for add_textbox, default: 200)"
            },
            textboxHeight = new
            {
                type = "number",
                description = "Textbox height in points (for add_textbox, default: 100)"
            },
            positionX = new
            {
                type = "number",
                description = "Horizontal position in points (for add_textbox, default: 100)"
            },
            positionY = new
            {
                type = "number",
                description = "Vertical position in points (for add_textbox, default: 100)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color hex (for add_textbox)"
            },
            borderColor = new
            {
                type = "string",
                description = "Border color hex (for add_textbox, set_textbox_border)"
            },
            borderWidth = new
            {
                type = "number",
                description = "Border width in points (for add_textbox, set_textbox_border, default: 1)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (for add_textbox, edit_textbox_content)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (for add_textbox, edit_textbox_content)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (for add_textbox, edit_textbox_content)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (for add_textbox, edit_textbox_content)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (for add_textbox, edit_textbox_content)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (for edit_textbox_content)"
            },
            color = new
            {
                type = "string",
                description = "Text color hex (for edit_textbox_content)"
            },
            textAlignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right (for add_textbox, default: left)",
                @enum = new[] { "left", "center", "right" }
            },
            textboxIndex = new
            {
                type = "number",
                description = "Textbox index (0-based, for edit_textbox_content, set_textbox_border)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, for set_textbox_border, default: 0, use -1 for all sections)"
            },
            appendText = new
            {
                type = "boolean",
                description = "Append text to existing content (for edit_textbox_content, default: false)"
            },
            clearFormatting = new
            {
                type = "boolean",
                description = "Clear existing formatting (for edit_textbox_content, default: false)"
            },
            borderVisible = new
            {
                type = "boolean",
                description = "Show border (for set_textbox_border, default: true)"
            },
            borderStyle = new
            {
                type = "string",
                description = "Border style (for set_textbox_border, default: solid)",
                @enum = new[] { "solid", "dash", "dot", "dashDot", "dashDotDot", "roundDot" }
            },
            includeContent = new
            {
                type = "boolean",
                description = "Include textbox content (for get_textboxes, default: true)"
            },
            chartType = new
            {
                type = "string",
                description = "Chart type: column, bar, line, pie, area, scatter, doughnut (for add_chart, default: column)",
                @enum = new[] { "column", "bar", "line", "pie", "area", "scatter", "doughnut" }
            },
            data = new
            {
                type = "array",
                description = "Chart data as 2D array (for add_chart)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            chartTitle = new
            {
                type = "string",
                description = "Chart title (for add_chart, optional)"
            },
            chartWidth = new
            {
                type = "number",
                description = "Chart width in points (for add_chart, default: 432)"
            },
            chartHeight = new
            {
                type = "number",
                description = "Chart height in points (for add_chart, default: 252)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert after (for add_chart, optional, use -1 for beginning)"
            },
            alignment = new
            {
                type = "string",
                description = "Chart alignment: left, center, right (for add_chart, default: left)",
                @enum = new[] { "left", "center", "right" }
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add_line" => await AddLine(arguments),
            "add_textbox" => await AddTextBox(arguments),
            "get_textboxes" => await GetTextboxes(arguments),
            "edit_textbox_content" => await EditTextBoxContent(arguments),
            "set_textbox_border" => await SetTextBoxBorder(arguments),
            "add_chart" => await AddChart(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddLine(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var location = arguments?["location"]?.GetValue<string>() ?? "body";
        var position = arguments?["position"]?.GetValue<string>() ?? "end";
        var lineStyle = arguments?["lineStyle"]?.GetValue<string>() ?? "shape";
        var lineWidth = arguments?["lineWidth"]?.GetValue<double?>() ?? 1.0;
        var lineColor = arguments?["lineColor"]?.GetValue<string>() ?? "000000";
        var width = arguments?["width"]?.GetValue<double?>();

        var doc = new Document(path);
        var section = doc.FirstSection;
        var calculatedWidth = width ?? (section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin);

        Node? targetNode = null;
        string locationDesc = "";

        switch (location.ToLower())
        {
            case "header":
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }
                targetNode = header;
                locationDesc = "header";
                break;

            case "footer":
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                    section.HeadersFooters.Add(footer);
                }
                targetNode = footer;
                locationDesc = "footer";
                break;

            case "body":
            default:
                targetNode = section.Body;
                locationDesc = "document body";
                break;
        }

        if (targetNode == null)
            throw new InvalidOperationException($"Could not access {location}");

        if (lineStyle == "shape")
        {
            var linePara = new Paragraph(doc);
            linePara.ParagraphFormat.SpaceBefore = 0;
            linePara.ParagraphFormat.SpaceAfter = 0;
            linePara.ParagraphFormat.LineSpacing = 1;
            linePara.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            
            var shape = new Shape(doc, ShapeType.Line);
            shape.Width = calculatedWidth;
            shape.Height = 0;
            shape.StrokeWeight = lineWidth;
            shape.StrokeColor = ParseColor(lineColor);
            shape.WrapType = WrapType.Inline;
            
            linePara.AppendChild(shape);

            if (position == "start")
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.PrependChild(linePara);
                else if (targetNode is Body body)
                    body.PrependChild(linePara);
            }
            else
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.AppendChild(linePara);
                else if (targetNode is Body body)
                    body.AppendChild(linePara);
            }
        }
        else
        {
            var linePara = new Paragraph(doc);
            linePara.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;
            linePara.ParagraphFormat.Borders.Bottom.LineWidth = lineWidth;
            linePara.ParagraphFormat.Borders.Bottom.Color = ParseColor(lineColor);

            if (position == "start")
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.PrependChild(linePara);
                else if (targetNode is Body body)
                    body.PrependChild(linePara);
            }
            else
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.AppendChild(linePara);
                else if (targetNode is Body body)
                    body.AppendChild(linePara);
            }
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Successfully inserted line in {locationDesc} at {position} position. Output: {outputPath}");
    }

    private async Task<string> AddTextBox(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var textboxWidth = arguments?["textboxWidth"]?.GetValue<double>() ?? 200;
        var textboxHeight = arguments?["textboxHeight"]?.GetValue<double>() ?? 100;
        var positionX = arguments?["positionX"]?.GetValue<double>() ?? 100;
        var positionY = arguments?["positionY"]?.GetValue<double>() ?? 100;
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();
        var borderColor = arguments?["borderColor"]?.GetValue<string>();
        var borderWidth = arguments?["borderWidth"]?.GetValue<double>() ?? 1;
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var textAlignment = arguments?["textAlignment"]?.GetValue<string>() ?? "left";

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = textboxWidth;
        textBox.Height = textboxHeight;
        textBox.Left = positionX;
        textBox.Top = positionY;
        textBox.WrapType = WrapType.None;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        if (!string.IsNullOrEmpty(backgroundColor))
        {
            textBox.Fill.Color = ParseColor(backgroundColor);
            textBox.Fill.Visible = true;
        }

        if (!string.IsNullOrEmpty(borderColor))
        {
            textBox.Stroke.Color = ParseColor(borderColor);
            textBox.Stroke.Weight = borderWidth;
            textBox.Stroke.Visible = true;
        }

        var para = new Paragraph(doc);
        var run = new Run(doc, text);

        if (!string.IsNullOrEmpty(fontNameAscii))
            run.Font.NameAscii = fontNameAscii;
        
        if (!string.IsNullOrEmpty(fontNameFarEast))
            run.Font.NameFarEast = fontNameFarEast;
        
        if (!string.IsNullOrEmpty(fontName))
        {
            if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                run.Font.Name = fontName;
            else
            {
                if (string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontName;
                if (string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontName;
            }
        }

        if (fontSize.HasValue)
            run.Font.Size = fontSize.Value;

        if (bold.HasValue)
            run.Font.Bold = bold.Value;

        para.ParagraphFormat.Alignment = textAlignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };

        para.AppendChild(run);
        textBox.AppendChild(para);
        builder.InsertNode(textBox);

        doc.Save(outputPath);
        return await Task.FromResult($"Successfully added textbox. Output: {outputPath}");
    }

    private async Task<string> GetTextboxes(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var includeContent = arguments?["includeContent"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox)
            .ToList();
        
        var result = new StringBuilder();
        result.AppendLine("=== Document Textboxes ===\n");
        result.AppendLine($"Total Textboxes: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("No textboxes found");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < shapes.Count; i++)
        {
            var textbox = shapes[i];
            result.AppendLine($"【Textbox {i}】");
            result.AppendLine($"Name: {textbox.Name ?? "(No name)"}");
            result.AppendLine($"Width: {textbox.Width} pt");
            result.AppendLine($"Height: {textbox.Height} pt");
            result.AppendLine($"Position: X={textbox.Left}, Y={textbox.Top}");
            
            if (includeContent)
            {
                var textboxText = textbox.GetText().Trim();
                if (!string.IsNullOrEmpty(textboxText))
                {
                    result.AppendLine($"Content:");
                    result.AppendLine($"  {textboxText.Replace("\n", "\n  ")}");
                }
                else
                    result.AppendLine($"Content: (empty)");
            }
            
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> EditTextBoxContent(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var textboxIndex = arguments?["textboxIndex"]?.GetValue<int>() ?? throw new ArgumentException("textboxIndex is required");
        var text = arguments?["text"]?.GetValue<string>();
        var appendText = arguments?["appendText"]?.GetValue<bool>() ?? false;
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var color = arguments?["color"]?.GetValue<string>();
        var clearFormatting = arguments?["clearFormatting"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        var textboxes = shapes.Cast<Shape>().Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        
        if (textboxIndex < 0 || textboxIndex >= textboxes.Count)
            throw new ArgumentException($"Textbox index {textboxIndex} out of range (total textboxes: {textboxes.Count})");
        
        var textbox = textboxes[textboxIndex];
        var paragraphs = textbox.GetChildNodes(NodeType.Paragraph, true);
        Paragraph para;
        
        if (paragraphs.Count == 0)
        {
            para = new Paragraph(doc);
            textbox.AppendChild(para);
        }
        else
            para = paragraphs[0] as Paragraph ?? throw new Exception("Cannot get textbox paragraph");
        
        if (text != null)
        {
            if (appendText && para.Runs.Count > 0)
            {
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
            }
            else
            {
                para.RemoveAllChildren();
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
            }
        }
        
        var runs = para.GetChildNodes(NodeType.Run, false);
        
        if (clearFormatting)
        {
            foreach (Run run in runs)
                run.Font.ClearFormatting();
        }
        
        bool hasFormatting = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) || 
                             !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue || 
                             bold.HasValue || italic.HasValue || !string.IsNullOrEmpty(color);
        
        if (hasFormatting)
        {
            foreach (Run run in runs)
            {
                if (!string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontNameAscii;
                
                if (!string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontNameFarEast;
                
                if (!string.IsNullOrEmpty(fontName))
                {
                    if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                        run.Font.Name = fontName;
                    else
                    {
                        if (string.IsNullOrEmpty(fontNameAscii))
                            run.Font.NameAscii = fontName;
                        if (string.IsNullOrEmpty(fontNameFarEast))
                            run.Font.NameFarEast = fontName;
                    }
                }
                
                if (fontSize.HasValue)
                    run.Font.Size = fontSize.Value;
                
                if (bold.HasValue)
                    run.Font.Bold = bold.Value;
                
                if (italic.HasValue)
                    run.Font.Italic = italic.Value;
                
                if (!string.IsNullOrEmpty(color))
                {
                    try
                    {
                        var colorStr = color.TrimStart('#');
                        if (colorStr.Length == 6)
                        {
                            int r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
                            int g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
                            int b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
                            run.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                        }
                    }
                    catch { }
                }
            }
        }
        
        doc.Save(outputPath);
        return await Task.FromResult($"Successfully edited textbox #{textboxIndex}. Output: {outputPath}");
    }

    private async Task<string> SetTextBoxBorder(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var textboxIndex = arguments?["textboxIndex"]?.GetValue<int>() ?? throw new ArgumentException("textboxIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var borderVisible = arguments?["borderVisible"]?.GetValue<bool>() ?? true;
        var borderColor = arguments?["borderColor"]?.GetValue<string>() ?? "000000";
        var borderWidth = arguments?["borderWidth"]?.GetValue<double>() ?? 1.0;

        var doc = new Document(path);
        
        List<Shape> allTextboxes = new List<Shape>();
        
        if (sectionIndex == -1)
        {
            foreach (Section section in doc.Sections)
            {
                var shapes = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                    .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
                allTextboxes.AddRange(shapes);
            }
        }
        else
        {
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
            
            var section = doc.Sections[sectionIndex];
            allTextboxes = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        }
        
        if (textboxIndex >= allTextboxes.Count)
            throw new ArgumentException($"Textbox index {textboxIndex} out of range (total textboxes: {allTextboxes.Count})");
        
        var textBox = allTextboxes[textboxIndex];
        var stroke = textBox.Stroke;
        
        stroke.Visible = borderVisible;
        
        if (borderVisible)
        {
            stroke.Color = ParseColor(borderColor);
            stroke.Weight = borderWidth;
        }
        
        doc.Save(outputPath);
        
        var borderDesc = borderVisible 
            ? $"Border: {borderWidth}pt, Color: {borderColor}"
            : "No border";
        
        return await Task.FromResult($"Successfully set textbox {textboxIndex} {borderDesc}. Output: {outputPath}");
    }

    private async Task<string> AddChart(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var chartType = arguments?["chartType"]?.GetValue<string>() ?? "column";
        var data = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required");
        var chartTitle = arguments?["chartTitle"]?.GetValue<string>();
        var chartWidth = arguments?["chartWidth"]?.GetValue<double>() ?? 432.0;
        var chartHeight = arguments?["chartHeight"]?.GetValue<double>() ?? 252.0;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var alignment = arguments?["alignment"]?.GetValue<string>() ?? "left";

        if (data.Count == 0)
            throw new ArgumentException("Chart data cannot be empty");

        var tableData = new List<List<string>>();
        foreach (var row in data)
        {
            if (row is JsonArray rowArray)
            {
                var rowData = new List<string>();
                foreach (var cell in rowArray)
                    rowData.Add(cell?.ToString() ?? "");
                tableData.Add(rowData);
            }
        }

        if (tableData.Count == 0)
            throw new ArgumentException("Cannot parse chart data");

        string tempExcelPath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.xlsx");
        try
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            
            for (int i = 0; i < tableData.Count; i++)
            {
                for (int j = 0; j < tableData[i].Count; j++)
                {
                    var cellValue = tableData[i][j];
                    if (double.TryParse(cellValue, out double numValue) && i > 0)
                        worksheet.Cells[i, j].PutValue(numValue);
                    else
                        worksheet.Cells[i, j].PutValue(cellValue);
                }
            }
            
            int maxCol = tableData.Max(r => r.Count);
            string dataRange = $"A1:{Convert.ToChar(64 + maxCol)}{tableData.Count}";
            
            var chartTypeEnum = chartType.ToLower() switch
            {
                "bar" => ChartType.Bar,
                "line" => ChartType.Line,
                "pie" => ChartType.Pie,
                "area" => ChartType.Area,
                "scatter" => ChartType.Scatter,
                "doughnut" => ChartType.Doughnut,
                _ => ChartType.Column
            };
            
            int chartIndex = worksheet.Charts.Add(chartTypeEnum, 0, tableData.Count + 2, 20, 10);
            var chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange(dataRange, true);
            
            if (!string.IsNullOrEmpty(chartTitle))
                chart.Title.Text = chartTitle;
            
            workbook.Save(tempExcelPath);
            
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.png");
            chart.ToImage(tempImagePath, Aspose.Cells.Drawing.ImageType.Png);
            
            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            
            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    if (paragraphs.Count > 0)
                    {
                        var firstPara = paragraphs[0] as Paragraph;
                        if (firstPara != null)
                            builder.MoveTo(firstPara);
                    }
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    var targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                    if (targetPara != null)
                        builder.MoveTo(targetPara);
                    else
                        throw new ArgumentException($"Cannot find paragraph at index {paragraphIndex.Value}");
                }
                else
                    throw new ArgumentException($"Paragraph index {paragraphIndex.Value} out of range (total paragraphs: {paragraphs.Count})");
            }
            else
                builder.MoveToDocumentEnd();
            
            builder.ParagraphFormat.Alignment = alignment.ToLower() switch
            {
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                _ => ParagraphAlignment.Left
            };
            
            var shape = builder.InsertImage(tempImagePath);
            shape.Width = chartWidth;
            shape.Height = chartHeight;
            shape.WrapType = WrapType.Inline;
            
            if (File.Exists(tempImagePath))
            {
                try { File.Delete(tempImagePath); } catch { }
            }
            
            doc.Save(outputPath);
            
            return await Task.FromResult($"Successfully added chart. Type: {chartType}, Data rows: {tableData.Count}. Output: {outputPath}");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error creating chart: {ex.Message}", ex);
        }
        finally
        {
            if (File.Exists(tempExcelPath))
            {
                try { File.Delete(tempExcelPath); } catch { }
            }
        }
    }

    private System.Drawing.Color ParseColor(string hexColor)
    {
        hexColor = hexColor.TrimStart('#');
        if (hexColor.Length == 6)
        {
            var r = Convert.ToByte(hexColor.Substring(0, 2), 16);
            var g = Convert.ToByte(hexColor.Substring(2, 2), 16);
            var b = Convert.ToByte(hexColor.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        return System.Drawing.Color.Black;
    }
}

