using System.ComponentModel;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using ImageType = Aspose.Cells.Drawing.ImageType;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing shapes (lines, textboxes, charts, etc.) in Word documents
/// </summary>
[McpServerToolType]
public class WordShapeTool
{
    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordShapeTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordShapeTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_shape")]
    [Description(
        @"Manage shapes in Word documents. Supports 9 operations: add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border, add_chart, add, get, delete.

Note: All position/size values are in points (1 point = 1/72 inch, 72 points = 1 inch).
Important: Textbox operations (get_textboxes, edit_textbox_content, set_textbox_border) use a separate textbox-only index system.
General shape operations (add, get, delete) use an index that includes ALL shapes (lines, rectangles, textboxes, images, etc.).

Usage examples:
- Add line: word_shape(operation='add_line', path='doc.docx')
- Add textbox: word_shape(operation='add_textbox', path='doc.docx', text='Textbox content', positionX=100, positionY=100, textboxWidth=200, textboxHeight=100)
- Get textboxes: word_shape(operation='get_textboxes', path='doc.docx')
- Edit textbox: word_shape(operation='edit_textbox_content', path='doc.docx', textboxIndex=0, text='Updated content')
- Set border: word_shape(operation='set_textbox_border', path='doc.docx', textboxIndex=0, borderColor='#FF0000', borderWidth=2)
- Add chart: word_shape(operation='add_chart', path='doc.docx', chartType='Column', data=[['A','B'],['1','2']])
- Add generic shape: word_shape(operation='add', path='doc.docx', shapeType='Rectangle', width=100, height=50)
- Get all shapes: word_shape(operation='get', path='doc.docx')
- Delete shape: word_shape(operation='delete', path='doc.docx', shapeIndex=0)")]
    public string Execute(
        [Description(
            "Operation: add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border, add_chart, add, get, delete")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Location: body, header, footer (for add_line, default: body)")]
        string location = "body",
        [Description("Position: start, end (for add_line, default: end)")]
        string position = "end",
        [Description("Line style: border, shape (for add_line, default: shape)")]
        string lineStyle = "shape",
        [Description("Line width in points (for add_line, default: 1.0)")]
        double lineWidth = 1.0,
        [Description("Line color hex (for add_line, default: 000000)")]
        string lineColor = "000000",
        [Description("Width/length in points (for add_line: line length; for add: shape width)")]
        double? width = null,
        [Description("Text content (for add_textbox, edit_textbox_content)")]
        string? text = null,
        [Description("Textbox width in points (for add_textbox, default: 200)")]
        double textboxWidth = 200,
        [Description("Textbox height in points (for add_textbox, default: 100)")]
        double textboxHeight = 100,
        [Description("Horizontal position in points (for add_textbox, default: 100)")]
        double positionX = 100,
        [Description("Vertical position in points (for add_textbox, default: 100)")]
        double positionY = 100,
        [Description("Background color hex (for add_textbox)")]
        string? backgroundColor = null,
        [Description("Border color hex (for add_textbox, set_textbox_border)")]
        string? borderColor = null,
        [Description("Border width in points (for add_textbox, set_textbox_border, default: 1)")]
        double borderWidth = 1,
        [Description("Font name (for add_textbox, edit_textbox_content)")]
        string? fontName = null,
        [Description("Font name for ASCII characters (for add_textbox, edit_textbox_content)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (for add_textbox, edit_textbox_content)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (for add_textbox, edit_textbox_content)")]
        double? fontSize = null,
        [Description("Bold text (for add_textbox, edit_textbox_content)")]
        bool? bold = null,
        [Description("Italic text (for edit_textbox_content)")]
        bool? italic = null,
        [Description("Text color hex (for edit_textbox_content)")]
        string? color = null,
        [Description("Text alignment: left, center, right (for add_textbox, default: left)")]
        string textAlignment = "left",
        [Description("Textbox index (0-based, textbox-only index for edit_textbox_content, set_textbox_border)")]
        int? textboxIndex = null,
        [Description("Shape index (0-based, global index including all shapes, for delete operation)")]
        int? shapeIndex = null,
        [Description("Shape type: rectangle, ellipse, roundrectangle, line (for add operation)")]
        string? shapeType = null,
        [Description("Shape height in points (for add operation)")]
        double? height = null,
        [Description("Shape X position in points (for add operation, default: 100)")]
        double x = 100,
        [Description("Shape Y position in points (for add operation, default: 100)")]
        double y = 100,
        [Description("Append text to existing content (for edit_textbox_content, default: false)")]
        bool appendText = false,
        [Description("Clear existing formatting (for edit_textbox_content, default: false)")]
        bool clearFormatting = false,
        [Description("Show border (for set_textbox_border, default: true)")]
        bool borderVisible = true,
        [Description(
            "Border style: solid, dash, dot, dashDot, dashDotDot, roundDot (for set_textbox_border, default: solid)")]
        string borderStyle = "solid",
        [Description("Include textbox content (for get_textboxes, default: true)")]
        bool includeContent = true,
        [Description("Chart type: column, bar, line, pie, area, scatter, doughnut (for add_chart, default: column)")]
        string chartType = "column",
        [Description("Chart data as 2D array (for add_chart)")]
        string[][]? data = null,
        [Description("Chart title (for add_chart, optional)")]
        string? chartTitle = null,
        [Description("Chart width in points (for add_chart, default: 432)")]
        double chartWidth = 432,
        [Description("Chart height in points (for add_chart, default: 252)")]
        double chartHeight = 252,
        [Description("Paragraph index to insert after (for add_chart, optional, use -1 for beginning)")]
        int? paragraphIndex = null,
        [Description("Chart alignment: left, center, right (for add_chart, default: left)")]
        string alignment = "left")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add_line" => AddLine(ctx, outputPath, location, position, lineStyle, lineWidth, lineColor, width),
            "add_textbox" => AddTextBox(ctx, outputPath, text, textboxWidth, textboxHeight, positionX, positionY,
                backgroundColor, borderColor, borderWidth, fontName, fontNameAscii, fontNameFarEast, fontSize, bold,
                textAlignment),
            "get_textboxes" => GetTextboxes(ctx, includeContent),
            "edit_textbox_content" => EditTextBoxContent(ctx, outputPath, textboxIndex, text, appendText, fontName,
                fontNameAscii, fontNameFarEast, fontSize, bold, italic, color, clearFormatting),
            "set_textbox_border" => SetTextBoxBorder(ctx, outputPath, textboxIndex, borderVisible, borderColor,
                borderWidth, borderStyle),
            "add_chart" => AddChart(ctx, outputPath, chartType, data, chartTitle, chartWidth, chartHeight,
                paragraphIndex, alignment),
            "add" => AddShape(ctx, outputPath, shapeType, width, height, x, y),
            "get" => GetShapes(ctx),
            "delete" => DeleteShape(ctx, outputPath, shapeIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a line shape to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="location">The location for the line: body, header, or footer.</param>
    /// <param name="position">The position: start or end.</param>
    /// <param name="lineStyle">The line style: border or shape.</param>
    /// <param name="lineWidth">The line width in points.</param>
    /// <param name="lineColor">The line color in hex format.</param>
    /// <param name="width">The optional line width/length in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="InvalidOperationException">Thrown when the target location cannot be accessed.</exception>
    private static string AddLine(DocumentContext<Document> ctx, string? outputPath, string location, string position,
        string lineStyle, double lineWidth, string lineColor, double? width)
    {
        var doc = ctx.Document;
        var section = doc.FirstSection;
        var calculatedWidth = width ?? section.PageSetup.PageWidth - section.PageSetup.LeftMargin -
            section.PageSetup.RightMargin;

        Node? targetNode;
        string locationDesc;

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

            default:
                targetNode = section.Body;
                locationDesc = "document body";
                break;
        }

        if (targetNode == null)
            throw new InvalidOperationException($"Could not access {location}");

        if (lineStyle == "shape")
        {
            var linePara = new Paragraph(doc)
            {
                ParagraphFormat =
                {
                    SpaceBefore = 0,
                    SpaceAfter = 0,
                    LineSpacing = 1,
                    LineSpacingRule = LineSpacingRule.Exactly
                }
            };

            var shape = new Shape(doc, ShapeType.Line)
            {
                Width = calculatedWidth,
                Height = 0,
                StrokeWeight = lineWidth,
                StrokeColor = ColorHelper.ParseColor(lineColor),
                WrapType = WrapType.Inline
            };

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
            linePara.ParagraphFormat.Borders.Bottom.Color = ColorHelper.ParseColor(lineColor);

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

        ctx.Save(outputPath);
        var result = $"Successfully inserted line in {locationDesc} at {position} position.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Adds a textbox to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The text content for the textbox.</param>
    /// <param name="textboxWidth">The textbox width in points.</param>
    /// <param name="textboxHeight">The textbox height in points.</param>
    /// <param name="positionX">The horizontal position in points.</param>
    /// <param name="positionY">The vertical position in points.</param>
    /// <param name="backgroundColor">The background color in hex format.</param>
    /// <param name="borderColor">The border color in hex format.</param>
    /// <param name="borderWidth">The border width in points.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text is bold.</param>
    /// <param name="textAlignment">The text alignment: left, center, or right.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is null or empty.</exception>
    private static string AddTextBox(DocumentContext<Document> ctx, string? outputPath, string? text,
        double textboxWidth, double textboxHeight, double positionX, double positionY, string? backgroundColor,
        string? borderColor, double borderWidth, string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, string textAlignment)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add_textbox operation");

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = textboxWidth,
            Height = textboxHeight,
            Left = positionX,
            Top = positionY,
            WrapType = WrapType.None,
            RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
            RelativeVerticalPosition = RelativeVerticalPosition.Page
        };

        if (!string.IsNullOrEmpty(backgroundColor))
        {
            textBox.Fill.Color = ColorHelper.ParseColor(backgroundColor);
            textBox.Fill.Visible = true;
        }

        if (!string.IsNullOrEmpty(borderColor))
        {
            textBox.Stroke.Color = ColorHelper.ParseColor(borderColor);
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
            {
                run.Font.Name = fontName;
            }
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

        ctx.Save(outputPath);
        var result = "Successfully added textbox.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets all textboxes from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="includeContent">Whether to include textbox content in the output.</param>
    /// <returns>A formatted string containing information about all textboxes.</returns>
    private static string GetTextboxes(DocumentContext<Document> ctx, bool includeContent)
    {
        var doc = ctx.Document;
        var shapes = FindAllTextboxes(doc);

        var result = new StringBuilder();
        result.AppendLine("=== Document Textboxes ===\n");
        result.AppendLine($"Total Textboxes: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("No textboxes found");
            return result.ToString();
        }

        for (var i = 0; i < shapes.Count; i++)
        {
            var textbox = shapes[i];
            result.AppendLine($"[Textbox {i}]");
            result.AppendLine($"Name: {textbox.Name ?? "(No name)"}");
            result.AppendLine($"Width: {textbox.Width} pt");
            result.AppendLine($"Height: {textbox.Height} pt");
            result.AppendLine($"Position: X={textbox.Left}, Y={textbox.Top}");

            if (includeContent)
            {
                var textboxText = textbox.GetText().Trim();
                if (!string.IsNullOrEmpty(textboxText))
                {
                    result.AppendLine("Content:");
                    result.AppendLine($"  {textboxText.Replace("\n", "\n  ")}");
                }
                else
                {
                    result.AppendLine("Content: (empty)");
                }
            }

            result.AppendLine();
        }

        return result.ToString();
    }

    /// <summary>
    ///     Edits the content of a textbox.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="textboxIndex">The zero-based index of the textbox to edit.</param>
    /// <param name="text">The new text content.</param>
    /// <param name="appendText">Whether to append text to existing content.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text is bold.</param>
    /// <param name="italic">Whether the text is italic.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="clearFormatting">Whether to clear existing formatting.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when textboxIndex is not provided or is out of range.</exception>
    private static string EditTextBoxContent(DocumentContext<Document> ctx, string? outputPath, int? textboxIndex,
        string? text, bool appendText, string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, bool? italic, string? color, bool clearFormatting)
    {
        if (!textboxIndex.HasValue)
            throw new ArgumentException("textboxIndex is required for edit_textbox_content operation");

        var doc = ctx.Document;
        var textboxes = FindAllTextboxes(doc);

        if (textboxIndex.Value < 0 || textboxIndex.Value >= textboxes.Count)
            throw new ArgumentException(
                $"Textbox index {textboxIndex.Value} out of range (total textboxes: {textboxes.Count})");

        var textbox = textboxes[textboxIndex.Value];

        // Use false to get only direct child paragraphs of the textbox (not recursive)
        var paragraphs = textbox.GetChildNodes(NodeType.Paragraph, false);
        Paragraph para;

        if (paragraphs.Count == 0)
        {
            // Create a new paragraph inside the textbox
            para = new Paragraph(doc);
            textbox.AppendChild(para);
        }
        else
        {
            // Use the first paragraph that is a direct child of the textbox
            para = paragraphs[0] as Paragraph ?? throw new Exception("Cannot get textbox paragraph");
        }

        // Ensure we're working only with content inside the textbox
        // Get runs that are direct children of the paragraph (which is inside the textbox)
        var runsCollection = para.GetChildNodes(NodeType.Run, false);
        var runs = runsCollection.Cast<Run>().ToList();

        if (text != null)
        {
            if (appendText && runsCollection.Count > 0)
            {
                // Append new run to existing content
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
            }
            else
            {
                // Clear existing content and set new text
                // Only remove direct children of the paragraph (runs inside textbox)
                para.RemoveAllChildren();
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
            }

            // Refresh runs list after modification
            runs = para.GetChildNodes(NodeType.Run, false).Cast<Run>().ToList();
        }

        if (clearFormatting)
            foreach (var run in runs)
                run.Font.ClearFormatting();

        var hasFormatting = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) ||
                            !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue ||
                            bold.HasValue || italic.HasValue || !string.IsNullOrEmpty(color);

        if (hasFormatting)
            foreach (var run in runs)
            {
                // Apply font settings using FontHelper
                FontHelper.Word.ApplyFontSettings(
                    run,
                    fontName,
                    fontNameAscii,
                    fontNameFarEast,
                    fontSize,
                    bold,
                    italic
                );

                // Handle color separately to throw exception on parse error
                if (!string.IsNullOrEmpty(color))
                    run.Font.Color = ColorHelper.ParseColor(color, true);
            }

        ctx.Save(outputPath);
        var result = $"Successfully edited textbox #{textboxIndex.Value}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets border properties for a textbox.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="textboxIndex">The zero-based index of the textbox.</param>
    /// <param name="borderVisible">Whether the border is visible.</param>
    /// <param name="borderColor">The border color in hex format.</param>
    /// <param name="borderWidth">The border width in points.</param>
    /// <param name="borderStyle">The border style (solid, dash, dot, etc.).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when textboxIndex is not provided or is out of range.</exception>
    private static string SetTextBoxBorder(DocumentContext<Document> ctx, string? outputPath, int? textboxIndex,
        bool borderVisible, string? borderColor, double borderWidth, string borderStyle)
    {
        if (!textboxIndex.HasValue)
            throw new ArgumentException("textboxIndex is required for set_textbox_border operation");

        var doc = ctx.Document;
        var allTextboxes = FindAllTextboxes(doc);

        if (textboxIndex.Value < 0 || textboxIndex.Value >= allTextboxes.Count)
            throw new ArgumentException(
                $"Textbox index {textboxIndex.Value} out of range (total textboxes: {allTextboxes.Count})");

        var textBox = allTextboxes[textboxIndex.Value];
        var stroke = textBox.Stroke;

        stroke.Visible = borderVisible;

        if (borderVisible)
        {
            stroke.Color = ColorHelper.ParseColor(borderColor ?? "000000");
            stroke.Weight = borderWidth;
            stroke.DashStyle = borderStyle.ToLower() switch
            {
                "dash" => DashStyle.Dash,
                "dot" => DashStyle.Dot,
                "dashdot" => DashStyle.DashDot,
                "dashdotdot" => DashStyle.LongDashDotDot,
                "rounddot" => DashStyle.ShortDot,
                _ => DashStyle.Solid
            };
        }

        ctx.Save(outputPath);

        var borderDesc = borderVisible
            ? $"Border: {borderWidth}pt, Color: {borderColor ?? "000000"}, Style: {borderStyle}"
            : "No border";

        var result = $"Successfully set textbox {textboxIndex.Value} {borderDesc}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Adds a chart to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="chartType">The chart type: column, bar, line, pie, area, scatter, or doughnut.</param>
    /// <param name="data">The chart data as a 2D array.</param>
    /// <param name="chartTitle">The optional chart title.</param>
    /// <param name="chartWidth">The chart width in points.</param>
    /// <param name="chartHeight">The chart height in points.</param>
    /// <param name="paragraphIndex">The optional paragraph index to insert after.</param>
    /// <param name="alignment">The chart alignment: left, center, or right.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when chart data is null or empty.</exception>
    /// <exception cref="InvalidOperationException">Thrown when an error occurs creating the chart.</exception>
    private static string AddChart(DocumentContext<Document> ctx, string? outputPath, string chartType,
        string[][]? data, string? chartTitle, double chartWidth, double chartHeight, int? paragraphIndex,
        string alignment)
    {
        if (data == null || data.Length == 0)
            throw new ArgumentException("Chart data cannot be empty");

        List<List<string>> tableData = [];
        foreach (var row in data)
        {
            List<string> rowData = [];
            foreach (var cell in row)
                rowData.Add(cell);
            tableData.Add(rowData);
        }

        if (tableData.Count == 0)
            throw new ArgumentException("Cannot parse chart data");

        var doc = ctx.Document;
        var tempExcelPath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.xlsx");
        try
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            for (var i = 0; i < tableData.Count; i++)
            for (var j = 0; j < tableData[i].Count; j++)
            {
                var cellValue = tableData[i][j];
                if (double.TryParse(cellValue, out var numValue) && i > 0)
                    worksheet.Cells[i, j].PutValue(numValue);
                else
                    worksheet.Cells[i, j].PutValue(cellValue);
            }

            var maxCol = tableData.Max(r => r.Count);
            var dataRange = $"A1:{Convert.ToChar(64 + maxCol)}{tableData.Count}";

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

            var chartIndex = worksheet.Charts.Add(chartTypeEnum, 0, tableData.Count + 2, 20, 10);
            var chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange(dataRange, true);

            if (!string.IsNullOrEmpty(chartTitle))
                chart.Title.Text = chartTitle;

            workbook.Save(tempExcelPath);

            var tempImagePath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.png");
            chart.ToImage(tempImagePath, ImageType.Png);

            var builder = new DocumentBuilder(doc);

            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    if (paragraphs.Count > 0)
                        if (paragraphs[0] is Paragraph firstPara)
                            builder.MoveTo(firstPara);
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                        builder.MoveTo(targetPara);
                    else
                        throw new ArgumentException($"Cannot find paragraph at index {paragraphIndex.Value}");
                }
                else
                {
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} out of range (total paragraphs: {paragraphs.Count})");
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
            }

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
                try
                {
                    File.Delete(tempImagePath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[WARN] Error deleting temp image file: {ex.Message}");
                }

            ctx.Save(outputPath);

            var result = $"Successfully added chart. Type: {chartType}, Data rows: {tableData.Count}.\n";
            result += ctx.GetOutputMessage(outputPath);
            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error creating chart: {ex.Message}", ex);
        }
        finally
        {
            if (File.Exists(tempExcelPath))
                try
                {
                    File.Delete(tempExcelPath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[WARN] Error deleting temp Excel file: {ex.Message}");
                }
        }
    }

    /// <summary>
    ///     Finds all textboxes in the document, searching in all sections' Body and HeaderFooter.
    ///     This ensures consistent textbox indexing across all operations.
    /// </summary>
    /// <param name="doc">The Word document to search.</param>
    /// <returns>A list of all textbox shapes found in the document.</returns>
    private static List<Shape> FindAllTextboxes(Document doc)
    {
        List<Shape> textboxes = [];
        foreach (var section in doc.Sections.Cast<Section>())
        {
            // Search in main body
            var bodyShapes = section.Body.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                .Where(s => s.ShapeType == ShapeType.TextBox);
            textboxes.AddRange(bodyShapes);

            // Search in headers and footers
            foreach (var header in section.HeadersFooters.Cast<HeaderFooter>())
            {
                var headerShapes = header.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                    .Where(s => s.ShapeType == ShapeType.TextBox);
                textboxes.AddRange(headerShapes);
            }
        }

        return textboxes;
    }

    /// <summary>
    ///     Adds a generic shape to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="shapeType">The shape type: rectangle, ellipse, roundrectangle, or line.</param>
    /// <param name="width">The shape width in points.</param>
    /// <param name="height">The shape height in points.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or shape type is unknown.</exception>
    private static string AddShape(DocumentContext<Document> ctx, string? outputPath, string? shapeType, double? width,
        double? height, double x, double y)
    {
        if (string.IsNullOrEmpty(shapeType))
            throw new ArgumentException("shapeType is required for add operation");
        if (!width.HasValue)
            throw new ArgumentException("width is required for add operation");
        if (!height.HasValue)
            throw new ArgumentException("height is required for add operation");

        var shapeTypeEnum = shapeType.ToLower() switch
        {
            "rectangle" => ShapeType.Rectangle,
            "ellipse" => ShapeType.Ellipse,
            "roundrectangle" => ShapeType.RoundRectangle,
            "line" => ShapeType.Line,
            _ => throw new ArgumentException($"Unknown shape type: {shapeType}")
        };

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertShape(shapeTypeEnum, width.Value, height.Value);
        shape.Left = x;
        shape.Top = y;

        ctx.Save(outputPath);
        var result = $"Successfully added {shapeType} shape.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets all shapes from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A formatted string containing information about all shapes.</returns>
    private static string GetShapes(DocumentContext<Document> ctx)
    {
        var doc = ctx.Document;
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

        var result = new StringBuilder();
        result.AppendLine("=== Document Shapes ===\n");
        result.AppendLine($"Total Shapes: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("No shapes found");
            return result.ToString();
        }

        for (var i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            result.AppendLine($"Shape {i}:");
            result.AppendLine($"  Type: {shape.ShapeType}");
            result.AppendLine($"  Name: {shape.Name ?? "(No name)"}");
            result.AppendLine($"  Size: {shape.Width} x {shape.Height} pt");
            result.AppendLine($"  Position: X={shape.Left}, Y={shape.Top}");
            result.AppendLine();
        }

        return result.ToString();
    }

    /// <summary>
    ///     Deletes a shape from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="shapeIndex">The zero-based index of the shape to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided or is out of range.</exception>
    private static string DeleteShape(DocumentContext<Document> ctx, string? outputPath, int? shapeIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        var doc = ctx.Document;
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

        if (shapeIndex.Value < 0 || shapeIndex.Value >= shapes.Count)
            throw new ArgumentException(
                $"Shape index {shapeIndex.Value} is out of range. Document has {shapes.Count} shapes.");

        shapes[shapeIndex.Value].Remove();
        ctx.Save(outputPath);

        var result = $"Successfully deleted shape at index {shapeIndex.Value}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }
}