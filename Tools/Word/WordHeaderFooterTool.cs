using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for header and footer operations in Word documents
///     Merges: WordSetHeaderTextTool, WordSetFooterTextTool, WordSetHeaderImageTool, WordSetFooterImageTool,
///     WordSetHeaderLineTool, WordSetFooterLineTool, WordSetHeaderTabStopsTool, WordSetFooterTabStopsTool,
///     WordSetHeaderFooterTool, WordGetHeadersFootersTool
/// </summary>
public class WordHeaderFooterTool : IAsposeTool
{
    public string Description =>
        @"Manage headers and footers in Word documents. Supports 10 operations: set_header_text, set_footer_text, set_header_image, set_footer_image, set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get.

Usage examples:
- Set header text: word_header_footer(operation='set_header_text', path='doc.docx', headerLeft='Left', headerCenter='Center', headerRight='Right')
- Set footer text: word_header_footer(operation='set_footer_text', path='doc.docx', footerLeft='Page', footerCenter='', footerRight='{PAGE}')
- Set header image: word_header_footer(operation='set_header_image', path='doc.docx', imagePath='logo.png')
- Get headers/footers: word_header_footer(operation='get', path='doc.docx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set_header_text': Set header text (required params: path)
- 'set_footer_text': Set footer text (required params: path)
- 'set_header_image': Set header image (required params: path, imagePath)
- 'set_footer_image': Set footer image (required params: path, imagePath)
- 'set_header_line': Set header line (required params: path)
- 'set_footer_line': Set footer line (required params: path)
- 'set_header_tabs': Set header tab stops (required params: path)
- 'set_footer_tabs': Set footer tab stops (required params: path)
- 'set_header_footer': Set header and footer together (required params: path)
- 'get': Get headers and footers info (required params: path)",
                @enum = new[]
                {
                    "set_header_text", "set_footer_text", "set_header_image", "set_footer_image", "set_header_line",
                    "set_footer_line", "set_header_tabs", "set_footer_tabs", "set_header_footer", "get"
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
                description = "Output file path (if not provided, overwrites input, for write operations)"
            },
            // Text parameters
            headerLeft = new
            {
                type = "string",
                description = "Header left section text (optional, for set_header_text operation)"
            },
            headerCenter = new
            {
                type = "string",
                description = "Header center section text (optional, for set_header_text operation)"
            },
            headerRight = new
            {
                type = "string",
                description = "Header right section text (optional, for set_header_text operation)"
            },
            footerLeft = new
            {
                type = "string",
                description = "Footer left section text (optional, for set_footer_text operation)"
            },
            footerCenter = new
            {
                type = "string",
                description = "Footer center section text (optional, for set_footer_text operation)"
            },
            footerRight = new
            {
                type = "string",
                description = "Footer right section text (optional, for set_footer_text operation)"
            },
            // Image parameters
            imagePath = new
            {
                type = "string",
                description = "Path to image file (required for set_header_image/set_footer_image operations)"
            },
            alignment = new
            {
                type = "string",
                description = "Image alignment: left, center, right (optional, default: left, for image operations)",
                @enum = new[] { "left", "center", "right" }
            },
            imageWidth = new
            {
                type = "number",
                description = "Image width in points (optional, default: 50, for image operations)"
            },
            imageHeight = new
            {
                type = "number",
                description =
                    "Image height in points (optional, maintains aspect ratio if not specified, for image operations)"
            },
            // Line parameters
            lineStyle = new
            {
                type = "string",
                description = "Line style: single, double, thick (optional, for line operations)",
                @enum = new[] { "single", "double", "thick" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Line width in points (optional, for line operations)"
            },
            // Tab stops parameters
            tabStops = new
            {
                type = "array",
                description = "Tab stops (optional, for tab operations)",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number" },
                        alignment = new
                            { type = "string", @enum = new[] { "left", "center", "right", "decimal", "bar" } },
                        leader = new { type = "string", @enum = new[] { "none", "dots", "dashes", "line" } }
                    }
                }
            },
            // Font parameters
            fontName = new
            {
                type = "string",
                description = "Font name (optional)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, optional)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, optional)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (optional)"
            },
            // Common parameters
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0, use -1 to apply to all sections)"
            },
            clearExisting = new
            {
                type = "boolean",
                description = "Clear existing content before setting new content (optional, default: true)"
            },
            removeExisting = new
            {
                type = "boolean",
                description =
                    "Remove existing images before adding new one (optional, default: true, for image operations)"
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
        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation switch
        {
            "set_header_text" => await SetHeaderTextAsync(arguments, path, outputPath),
            "set_footer_text" => await SetFooterTextAsync(arguments, path, outputPath),
            "set_header_image" => await SetHeaderImageAsync(arguments, path, outputPath),
            "set_footer_image" => await SetFooterImageAsync(arguments, path, outputPath),
            "set_header_line" => await SetHeaderLineAsync(arguments, path, outputPath),
            "set_footer_line" => await SetFooterLineAsync(arguments, path, outputPath),
            "set_header_tabs" => await SetHeaderTabStopsAsync(arguments, path, outputPath),
            "set_footer_tabs" => await SetFooterTabStopsAsync(arguments, path, outputPath),
            "set_header_footer" => await SetHeaderFooterAsync(arguments, path, outputPath),
            "get" => await GetHeadersFootersAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets header text for the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing optional headerLeft, headerCenter, headerRight, sectionIndex,
    ///     differentFirstPage, differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetHeaderTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var headerLeft = ArgumentHelper.GetStringNullable(arguments, "headerLeft");
            var headerCenter = ArgumentHelper.GetStringNullable(arguments, "headerCenter");
            var headerRight = ArgumentHelper.GetStringNullable(arguments, "headerRight");
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var clearExisting = ArgumentHelper.GetBool(arguments, "clearExisting", true);

            var doc = new Document(path);

            var hasContent = !string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) ||
                             !string.IsNullOrEmpty(headerRight);
            if (!hasContent)
                return "Warning: No header text content provided";

            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }
                else if (clearExisting)
                {
                    header.RemoveAllChildren();
                }

                if (hasContent)
                {
                    header.RemoveAllChildren();
                    var headerPara = new Paragraph(doc);
                    header.AppendChild(headerPara);

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

            doc.Save(outputPath);

            var contentParts = new List<string>();
            if (!string.IsNullOrEmpty(headerLeft)) contentParts.Add("left");
            if (!string.IsNullOrEmpty(headerCenter)) contentParts.Add("center");
            if (!string.IsNullOrEmpty(headerRight)) contentParts.Add("right");

            var contentDesc = string.Join(", ", contentParts);
            var sectionsDesc = sectionIndex == -1 ? "all sections" : $"section {sectionIndex}";

            return $"Header text set successfully ({contentDesc}) in {sectionsDesc}";
        });
    }

    /// <summary>
    ///     Sets footer text for the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing optional footerLeft, footerCenter, footerRight, sectionIndex,
    ///     differentFirstPage, differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetFooterTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var footerLeft = ArgumentHelper.GetStringNullable(arguments, "footerLeft");
            var footerCenter = ArgumentHelper.GetStringNullable(arguments, "footerCenter");
            var footerRight = ArgumentHelper.GetStringNullable(arguments, "footerRight");
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var clearExisting = ArgumentHelper.GetBool(arguments, "clearExisting", true);

            var doc = new Document(path);

            var hasContent = !string.IsNullOrEmpty(footerLeft) || !string.IsNullOrEmpty(footerCenter) ||
                             !string.IsNullOrEmpty(footerRight);
            if (!hasContent)
                return "Warning: No footer text content provided";

            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                    section.HeadersFooters.Add(footer);
                }
                else if (clearExisting)
                {
                    footer.RemoveAllChildren();
                }

                if (hasContent)
                {
                    footer.RemoveAllChildren();
                    var footerPara = new Paragraph(doc);
                    footer.AppendChild(footerPara);

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

            doc.Save(outputPath);

            var contentParts = new List<string>();
            if (!string.IsNullOrEmpty(footerLeft)) contentParts.Add("left");
            if (!string.IsNullOrEmpty(footerCenter)) contentParts.Add("center");
            if (!string.IsNullOrEmpty(footerRight)) contentParts.Add("right");

            var contentDesc = string.Join(", ", contentParts);
            var sectionsDesc = sectionIndex == -1 ? "all sections" : $"section {sectionIndex}";

            return $"Footer text set successfully ({contentDesc}) in {sectionsDesc}";
        });
    }

    /// <summary>
    ///     Sets header image for the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing imagePath, optional sectionIndex, differentFirstPage,
    ///     differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetHeaderImageAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
            var alignment = ArgumentHelper.GetString(arguments, "alignment", "left");
            var imageWidth = ArgumentHelper.GetDoubleNullable(arguments, "imageWidth");
            var imageHeight = ArgumentHelper.GetDoubleNullable(arguments, "imageHeight");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var removeExisting = ArgumentHelper.GetBool(arguments, "removeExisting", true);

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            var doc = new Document(path);
            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }

                if (removeExisting)
                {
                    var existingShapes = header.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                        .Where(s => s.HasImage).ToList();
                    foreach (var existingShape in existingShapes) existingShape.Remove();
                }

                // Create a new paragraph for the image (similar to SetHeaderTextAsync)
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

                // Insert image (InsertImage creates inline image by default)
                var shape = builder.InsertImage(imagePath);
                if (imageWidth.HasValue) shape.Width = imageWidth.Value;
                if (imageHeight.HasValue) shape.Height = imageHeight.Value;

                // Ensure the paragraph containing the shape has correct alignment
                headerPara.ParagraphFormat.Alignment = paraAlignment;
            }

            doc.Save(outputPath);
            return $"Header image set: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets footer image for the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing imagePath, optional sectionIndex, differentFirstPage,
    ///     differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetFooterImageAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
            var alignment = ArgumentHelper.GetString(arguments, "alignment", "left");
            var imageWidth = ArgumentHelper.GetDoubleNullable(arguments, "imageWidth");
            var imageHeight = ArgumentHelper.GetDoubleNullable(arguments, "imageHeight");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var removeExisting = ArgumentHelper.GetBool(arguments, "removeExisting", true);

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            var doc = new Document(path);
            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                    section.HeadersFooters.Add(footer);
                }

                if (removeExisting)
                {
                    var existingShapes = footer.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                        .Where(s => s.HasImage).ToList();
                    foreach (var existingShape in existingShapes) existingShape.Remove();
                }

                // Create a new paragraph for the image (similar to SetFooterTextAsync)
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

                // Insert image (InsertImage creates inline image by default)
                var shape = builder.InsertImage(imagePath);
                if (imageWidth.HasValue) shape.Width = imageWidth.Value;
                if (imageHeight.HasValue) shape.Height = imageHeight.Value;

                // Ensure the paragraph containing the shape has correct alignment
                footerPara.ParagraphFormat.Alignment = paraAlignment;
            }

            doc.Save(outputPath);
            return $"Footer image set: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets header line (border) for the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing optional color, width, style, sectionIndex, differentFirstPage,
    ///     differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetHeaderLineAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var lineStyle = ArgumentHelper.GetString(arguments, "lineStyle", "single");
            var lineWidth = ArgumentHelper.GetDoubleNullable(arguments, "lineWidth");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
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

            doc.Save(outputPath);
            return $"Header line set: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets footer line (border) for the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing optional color, width, style, sectionIndex, differentFirstPage,
    ///     differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetFooterLineAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var lineStyle = ArgumentHelper.GetString(arguments, "lineStyle", "single");
            var lineWidth = ArgumentHelper.GetDoubleNullable(arguments, "lineWidth");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
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

            doc.Save(outputPath);
            return $"Footer line set: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets header tab stops
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing tabStops array, optional sectionIndex, differentFirstPage,
    ///     differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetHeaderTabStopsAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var tabStops = ArgumentHelper.GetArray(arguments, "tabStops", false);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
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

            doc.Save(outputPath);
            return $"Header tab stops set: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets footer tab stops
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing tabStops array, optional sectionIndex, differentFirstPage,
    ///     differentOddEven
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetFooterTabStopsAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var tabStops = ArgumentHelper.GetArray(arguments, "tabStops", false);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

            foreach (var section in sections)
            {
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
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

            doc.Save(outputPath);
            return $"Footer tab stops set: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets header and footer together
    /// </summary>
    /// <param name="arguments">JSON arguments containing header/footer properties</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetHeaderFooterAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            // Combined operation to set both header and footer
            SetHeaderTextAsync(arguments, path, outputPath).Wait();
            SetFooterTextAsync(arguments, path, outputPath).Wait();
            return $"Header and footer set: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all headers and footers from the document
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all headers and footers</returns>
    private Task<string> GetHeadersFootersAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);
            doc.UpdateFields();

            var result = new StringBuilder();

            result.AppendLine("=== Document Header and Footer Information ===\n");
            result.AppendLine($"Total sections: {doc.Sections.Count}\n");

            var sections = sectionIndex.HasValue
                ? [doc.Sections[sectionIndex.Value]]
                : doc.Sections.Cast<Section>().ToArray();

            if (sectionIndex.HasValue && (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
                throw new ArgumentException(
                    $"Section index {sectionIndex.Value} is out of range (document has {doc.Sections.Count} sections)");

            for (var i = 0; i < sections.Length; i++)
            {
                var section = sections[i];
                var actualIndex = sectionIndex ?? i;

                result.AppendLine($"[Section {actualIndex}]");

                result.AppendLine("Headers:");
                var headerTypes = new[]
                {
                    (HeaderFooterType.HeaderPrimary, "Primary header"),
                    (HeaderFooterType.HeaderFirst, "First page header"),
                    (HeaderFooterType.HeaderEven, "Even page header")
                };

                var hasHeader = false;
                foreach (var (type, name) in headerTypes)
                {
                    var header = section.HeadersFooters[type];
                    if (header != null)
                    {
                        var headerText = header.ToString(SaveFormat.Text).Trim();
                        if (!string.IsNullOrEmpty(headerText))
                        {
                            result.AppendLine($"  {name}:");
                            result.AppendLine($"    {headerText.Replace("\n", "\n    ")}");
                            hasHeader = true;
                        }
                    }
                }

                if (!hasHeader) result.AppendLine("  (No header)");

                result.AppendLine();

                result.AppendLine("Footers:");
                var footerTypes = new[]
                {
                    (HeaderFooterType.FooterPrimary, "Primary footer"),
                    (HeaderFooterType.FooterFirst, "First page footer"),
                    (HeaderFooterType.FooterEven, "Even page footer")
                };

                var hasFooter = false;
                foreach (var (type, name) in footerTypes)
                {
                    var footer = section.HeadersFooters[type];
                    if (footer != null)
                    {
                        var footerText = footer.ToString(SaveFormat.Text).Trim();
                        if (!string.IsNullOrEmpty(footerText))
                        {
                            result.AppendLine($"  {name}:");
                            result.AppendLine($"    {footerText.Replace("\n", "\n    ")}");
                            hasFooter = true;
                        }
                    }
                }

                if (!hasFooter) result.AppendLine("  (No footer)");

                if (i < sections.Length - 1) result.AppendLine();
            }

            return result.ToString();
        });
    }

    /// <summary>
    ///     Inserts text or field code. If text contains field codes like {PAGE}, {DATE}, {NUMPAGES}, etc.,
    ///     they will be inserted as actual fields instead of plain text.
    /// </summary>
    private void InsertTextOrField(DocumentBuilder builder, string text, string? fontName, string? fontNameAscii,
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
    ///     Restores DocumentBuilder font to Normal style
    /// </summary>
    private static void RestoreNormalFont(DocumentBuilder builder)
    {
        builder.Font.Name = builder.Document.Styles[StyleIdentifier.Normal].Font.Name;
        builder.Font.Size = builder.Document.Styles[StyleIdentifier.Normal].Font.Size;
    }
}