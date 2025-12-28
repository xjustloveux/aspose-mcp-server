using System.Text;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;
using ImageType = Aspose.Cells.Drawing.ImageType;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing shapes (lines, textboxes, charts, etc.) in Word documents
/// </summary>
public class WordShapeTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
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
- Delete shape: word_shape(operation='delete', path='doc.docx', shapeIndex=0)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add_line': Add a line shape (required params: path)
- 'add_textbox': Add a textbox (required params: path, text)
- 'get_textboxes': Get all textboxes with textbox-only index (required params: path)
- 'edit_textbox_content': Edit textbox content using textbox-only index (required params: path, textboxIndex)
- 'set_textbox_border': Set textbox border using textbox-only index (required params: path, textboxIndex)
- 'add_chart': Add a chart (required params: path, data)
- 'add': Add a generic shape (required params: path, shapeType, width, height)
- 'get': Get all shapes with global shape index (required params: path)
- 'delete': Delete a shape using global shape index (required params: path, shapeIndex)",
                @enum = new[]
                {
                    "add_line", "add_textbox", "get_textboxes", "edit_textbox_content", "set_textbox_border",
                    "add_chart", "add", "get", "delete"
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
                description =
                    "Width/length in points (for add_line: line length; for add: shape width. 1 point = 1/72 inch)"
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
                description = "Textbox index (0-based, textbox-only index for edit_textbox_content, set_textbox_border)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, global index including all shapes, for delete operation)"
            },
            shapeType = new
            {
                type = "string",
                description = "Shape type (for add operation)",
                @enum = new[] { "rectangle", "ellipse", "roundrectangle", "line" }
            },
            height = new
            {
                type = "number",
                description = "Shape height in points (for add operation, 1 point = 1/72 inch)"
            },
            x = new
            {
                type = "number",
                description = "Shape X position in points (for add operation, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Shape Y position in points (for add operation, default: 100)"
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
                description =
                    "Chart type: column, bar, line, pie, area, scatter, doughnut (for add_chart, default: column)",
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "add_line" => await AddLine(path, outputPath, arguments),
            "add_textbox" => await AddTextBox(path, outputPath, arguments),
            "get_textboxes" => await GetTextboxes(path, arguments),
            "edit_textbox_content" => await EditTextBoxContent(path, outputPath, arguments),
            "set_textbox_border" => await SetTextBoxBorder(path, outputPath, arguments),
            "add_chart" => await AddChart(path, outputPath, arguments),
            "add" => await AddShape(path, outputPath, arguments),
            "get" => await GetShapes(path),
            "delete" => await DeleteShape(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a line shape to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional location, position, lineStyle, lineWidth, lineColor, width</param>
    /// <returns>Success message with line details</returns>
    private Task<string> AddLine(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var location = ArgumentHelper.GetString(arguments, "location", "body");
            var position = ArgumentHelper.GetString(arguments, "position", "end");
            var lineStyle = ArgumentHelper.GetString(arguments, "lineStyle", "shape");
            var lineWidth = ArgumentHelper.GetDouble(arguments, "lineWidth", 1.0);
            var lineColor = ArgumentHelper.GetString(arguments, "lineColor", "000000");
            var width = ArgumentHelper.GetDoubleNullable(arguments, "width");

            var doc = new Document(path);
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

            doc.Save(outputPath);
            return $"Successfully inserted line in {locationDesc} at {position} position. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds a textbox to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, optional textboxWidth, textboxHeight, positionX, positionY</param>
    /// <returns>Success message with textbox details</returns>
    private Task<string> AddTextBox(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var text = ArgumentHelper.GetString(arguments, "text");
            var textboxWidth = ArgumentHelper.GetDouble(arguments, "textboxWidth", "textboxWidth", false, 200);
            var textboxHeight = ArgumentHelper.GetDouble(arguments, "textboxHeight", "textboxHeight", false, 100);
            var positionX = ArgumentHelper.GetDouble(arguments, "positionX", "positionX", false, 100);
            var positionY = ArgumentHelper.GetDouble(arguments, "positionY", "positionY", false, 100);
            var backgroundColor = ArgumentHelper.GetStringNullable(arguments, "backgroundColor");
            var borderColor = ArgumentHelper.GetStringNullable(arguments, "borderColor");
            var borderWidth = ArgumentHelper.GetDouble(arguments, "borderWidth", "borderWidth", false, 1);
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var textAlignment = ArgumentHelper.GetString(arguments, "textAlignment", "left");

            var doc = new Document(path);
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

            doc.Save(outputPath);
            return $"Successfully added textbox. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all textboxes from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing optional includeContent</param>
    /// <returns>Formatted string with all textboxes</returns>
    private Task<string> GetTextboxes(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var includeContent = ArgumentHelper.GetBool(arguments, "includeContent", true);

            var doc = new Document(path);
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
                result.AppendLine($"�iTextbox {i}�j");
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
        });
    }

    /// <summary>
    ///     Edits the content of a textbox
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing textboxIndex, text, optional formatting options</param>
    /// <returns>Success message with updated textbox details</returns>
    private Task<string> EditTextBoxContent(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var textboxIndex = ArgumentHelper.GetInt(arguments, "textboxIndex");
            var text = ArgumentHelper.GetStringNullable(arguments, "text");
            var appendText = ArgumentHelper.GetBool(arguments, "appendText", false);
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
            var color = ArgumentHelper.GetStringNullable(arguments, "color");
            var clearFormatting = ArgumentHelper.GetBool(arguments, "clearFormatting", false);

            var doc = new Document(path);
            var textboxes = FindAllTextboxes(doc);

            if (textboxIndex < 0 || textboxIndex >= textboxes.Count)
                throw new ArgumentException(
                    $"Textbox index {textboxIndex} out of range (total textboxes: {textboxes.Count})");

            var textbox = textboxes[textboxIndex];

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

            doc.Save(outputPath);
            return $"Successfully edited textbox #{textboxIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets border properties for a textbox
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing textboxIndex, optional borderColor, borderWidth</param>
    /// <returns>Success message</returns>
    private Task<string> SetTextBoxBorder(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var textboxIndex = ArgumentHelper.GetInt(arguments, "textboxIndex");
            var borderVisible = ArgumentHelper.GetBool(arguments, "borderVisible", true);
            var borderColor = ArgumentHelper.GetString(arguments, "borderColor", "000000");
            var borderWidth = ArgumentHelper.GetDouble(arguments, "borderWidth", "borderWidth", false, 1.0);

            var doc = new Document(path);
            var allTextboxes = FindAllTextboxes(doc);

            if (textboxIndex < 0 || textboxIndex >= allTextboxes.Count)
                throw new ArgumentException(
                    $"Textbox index {textboxIndex} out of range (total textboxes: {allTextboxes.Count})");

            var textBox = allTextboxes[textboxIndex];
            var stroke = textBox.Stroke;

            stroke.Visible = borderVisible;

            if (borderVisible)
            {
                stroke.Color = ColorHelper.ParseColor(borderColor);
                stroke.Weight = borderWidth;
            }

            doc.Save(outputPath);

            var borderDesc = borderVisible
                ? $"Border: {borderWidth}pt, Color: {borderColor}"
                : "No border";

            return $"Successfully set textbox {textboxIndex} {borderDesc}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds a chart to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing chartType, data, optional chartTitle, chartWidth, chartHeight</param>
    /// <returns>Success message with chart details</returns>
    private Task<string> AddChart(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(async () =>
        {
            var chartType = ArgumentHelper.GetString(arguments, "chartType", "column");
            var data = ArgumentHelper.GetArray(arguments, "data");
            var chartTitle = ArgumentHelper.GetStringNullable(arguments, "chartTitle");
            var chartWidth = ArgumentHelper.GetDouble(arguments, "chartWidth", "chartWidth", false, 432.0);
            var chartHeight = ArgumentHelper.GetDouble(arguments, "chartHeight", "chartHeight", false, 252.0);
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var alignment = ArgumentHelper.GetString(arguments, "alignment", "left");

            if (data.Count == 0)
                throw new ArgumentException("Chart data cannot be empty");

            var tableData = new List<List<string>>();
            foreach (var row in data)
                if (row is JsonArray rowArray)
                {
                    var rowData = new List<string>();
                    foreach (var cell in rowArray)
                        rowData.Add(cell?.ToString() ?? "");
                    tableData.Add(rowData);
                }

            if (tableData.Count == 0)
                throw new ArgumentException("Cannot parse chart data");

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

                var doc = new Document(path);
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
                        await Console.Error.WriteLineAsync($"[WARN] Error deleting temp image file: {ex.Message}");
                    }

                doc.Save(outputPath);

                return
                    $"Successfully added chart. Type: {chartType}, Data rows: {tableData.Count}. Output: {outputPath}";
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
                        await Console.Error.WriteLineAsync($"[WARN] Error deleting temp Excel file: {ex.Message}");
                    }
            }
        });
    }

    /// <summary>
    ///     Find all textboxes in the document, searching in all sections' Body and HeaderFooter.
    ///     This ensures consistent textbox indexing across all operations.
    /// </summary>
    private List<Shape> FindAllTextboxes(Document doc)
    {
        var textboxes = new List<Shape>();
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
    ///     Adds a generic shape to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing shapeType, width, height, optional x, y</param>
    /// <returns>Success message with shape details</returns>
    private Task<string> AddShape(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeTypeStr = ArgumentHelper.GetString(arguments, "shapeType");
            var width = ArgumentHelper.GetDouble(arguments, "width");
            var height = ArgumentHelper.GetDouble(arguments, "height");
            var x = ArgumentHelper.GetDoubleNullable(arguments, "x") ?? 100;
            var y = ArgumentHelper.GetDoubleNullable(arguments, "y") ?? 100;

            var shapeType = shapeTypeStr.ToLower() switch
            {
                "rectangle" => ShapeType.Rectangle,
                "ellipse" => ShapeType.Ellipse,
                "roundrectangle" => ShapeType.RoundRectangle,
                "line" => ShapeType.Line,
                _ => throw new ArgumentException($"Unknown shape type: {shapeTypeStr}")
            };

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            var shape = builder.InsertShape(shapeType, width, height);
            shape.Left = x;
            shape.Top = y;

            doc.Save(outputPath);
            return $"Successfully added {shapeTypeStr} shape. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all shapes from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <returns>Formatted string with all shapes</returns>
    private Task<string> GetShapes(string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
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
        });
    }

    /// <summary>
    ///     Deletes a shape from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing shapeIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteShape(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            var doc = new Document(path);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

            if (shapeIndex < 0 || shapeIndex >= shapes.Count)
                throw new ArgumentException(
                    $"Shape index {shapeIndex} is out of range. Document has {shapes.Count} shapes.");

            shapes[shapeIndex].Remove();
            doc.Save(outputPath);

            return $"Successfully deleted shape at index {shapeIndex}. Output: {outputPath}";
        });
    }
}