using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetHeaderFooterTool : IAsposeTool
{
    public string Description => "Set header and footer content in a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            headerText = new
            {
                type = "string",
                description = "Header text (optional, for simple single-line header)"
            },
            footerText = new
            {
                type = "string",
                description = "Footer text (optional, for simple single-line footer)"
            },
            headerLeft = new
            {
                type = "string",
                description = "Header left section text (optional)"
            },
            headerCenter = new
            {
                type = "string",
                description = "Header center section text (optional)"
            },
            headerRight = new
            {
                type = "string",
                description = "Header right section text (optional)"
            },
            footerLeft = new
            {
                type = "string",
                description = "Footer left section text (optional)"
            },
            footerCenter = new
            {
                type = "string",
                description = "Footer center section text (optional)"
            },
            footerRight = new
            {
                type = "string",
                description = "Footer right section text (optional)"
            },
            headerAlignment = new
            {
                type = "string",
                description = "Header alignment when using headerText: left, center, right (default: center)",
                @enum = new[] { "left", "center", "right" }
            },
            footerAlignment = new
            {
                type = "string",
                description = "Footer alignment when using footerText: left, center, right (default: center)",
                @enum = new[] { "left", "center", "right" }
            },
            headerFontName = new
            {
                type = "string",
                description = "Header font name (e.g., '微軟正黑體', 'Arial'). If headerFontNameAscii and headerFontNameFarEast are provided, this will be used as fallback."
            },
            headerFontNameAscii = new
            {
                type = "string",
                description = "Header font name for ASCII characters (English, e.g., 'Times New Roman')"
            },
            headerFontNameFarEast = new
            {
                type = "string",
                description = "Header font name for Far East characters (Chinese/Japanese/Korean, e.g., '標楷體')"
            },
            headerFontSize = new
            {
                type = "number",
                description = "Header font size in points (e.g., 10)"
            },
            footerFontName = new
            {
                type = "string",
                description = "Footer font name (e.g., '微軟正黑體', 'Arial'). If footerFontNameAscii and footerFontNameFarEast are provided, this will be used as fallback."
            },
            footerFontNameAscii = new
            {
                type = "string",
                description = "Footer font name for ASCII characters (English, e.g., 'Times New Roman')"
            },
            footerFontNameFarEast = new
            {
                type = "string",
                description = "Footer font name for Far East characters (Chinese/Japanese/Korean, e.g., '標楷體')"
            },
            footerFontSize = new
            {
                type = "number",
                description = "Footer font size in points (e.g., 9)"
            },
            includePageNumber = new
            {
                type = "boolean",
                description = "Include page number in footer (default: false)"
            },
            pageNumberFormat = new
            {
                type = "string",
                description = "Page number format: simple (Page X), total (Page X of Y), chinese (第X頁 共Y頁), custom (use footerTemplate) (default: simple)",
                @enum = new[] { "simple", "total", "chinese", "custom" }
            },
            footerTemplate = new
            {
                type = "string",
                description = "Custom footer template with placeholders: {PAGE} (current page), {NUMPAGES} (total pages), {SECTION} (section number). Example: 'B03-{PAGE}' => 'B03-1', 'B03-2', etc. Requires pageNumberFormat='custom'."
            },
            headerBorder = new
            {
                type = "boolean",
                description = "Show border line below header (default: true)"
            },
            footerBorder = new
            {
                type = "boolean",
                description = "Show border line above footer (default: true)"
            },
            headerLineStyle = new
            {
                type = "string",
                description = "Header line style: 'border' (paragraph border, default) or 'shape' (graphic line)",
                @enum = new[] { "border", "shape" }
            },
            headerLineWidth = new
            {
                type = "number",
                description = "Header line width in points (default: 0.5 for border, 1.0 for shape)"
            },
            headerLineColor = new
            {
                type = "string",
                description = "Header line color in hex (e.g., '000000' for black, default: '000000')"
            },
            headerTabStops = new
            {
                type = "array",
                description = "Custom tab stops for header (optional). Example: [{\"position\": 70.90, \"alignment\": \"Left\"}, {\"position\": 541.45, \"alignment\": \"Right\"}]. If not specified, auto-calculates based on headerLeft/Center/Right.",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number", description = "Tab stop position in points" },
                        alignment = new { type = "string", description = "Tab alignment: Left, Center, Right, Decimal, Bar", @enum = new[] { "Left", "Center", "Right", "Decimal", "Bar" } },
                        leader = new { type = "string", description = "Tab leader: None, Dots, Dashes, Line, Heavy, MiddleDot (default: None)", @enum = new[] { "None", "Dots", "Dashes", "Line", "Heavy", "MiddleDot" } }
                    }
                }
            },
            headerImagePath = new
            {
                type = "string",
                description = "Path to image file for header (optional)"
            },
            headerImageWidth = new
            {
                type = "number",
                description = "Header image width in points (default: 50)"
            },
            headerImageHeight = new
            {
                type = "number",
                description = "Header image height in points (default: auto-calculate from width)"
            },
            headerImageAlignment = new
            {
                type = "string",
                description = "Header image alignment: left, center, right (default: left)",
                @enum = new[] { "left", "center", "right" }
            },
            footerImagePath = new
            {
                type = "string",
                description = "Path to image file for footer (optional)"
            },
            footerImageWidth = new
            {
                type = "number",
                description = "Footer image width in points (default: 20)"
            },
            footerImageHeight = new
            {
                type = "number",
                description = "Footer image height in points (default: auto-calculate from width)"
            },
            footerImageAlignment = new
            {
                type = "string",
                description = "Footer image alignment: left, center, right (default: left)",
                @enum = new[] { "left", "center", "right" }
            },
            footerLineStyle = new
            {
                type = "string",
                description = "Footer line style: 'border' (paragraph border, default) or 'shape' (graphic line)",
                @enum = new[] { "border", "shape" }
            },
            footerLineWidth = new
            {
                type = "number",
                description = "Footer line width in points (default: 0.5 for border, 1.0 for shape)"
            },
            footerLineColor = new
            {
                type = "string",
                description = "Footer line color in hex (e.g., '000000' for black, default: '000000')"
            },
            footerLinePosition = new
            {
                type = "string",
                description = "Footer line position: 'above' (above footer content, default) or 'below' (below footer content)",
                @enum = new[] { "above", "below" }
            },
            footerTabStops = new
            {
                type = "array",
                description = "Custom tab stops for footer (optional). Example: [{\"position\": 70.90, \"alignment\": \"Left\"}, {\"position\": 541.45, \"alignment\": \"Right\"}]. Set to empty array [] to remove all tab stops. If not specified, auto-calculates based on footerLeft/Center/Right.",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number", description = "Tab stop position in points" },
                        alignment = new { type = "string", description = "Tab alignment: Left, Center, Right, Decimal, Bar", @enum = new[] { "Left", "Center", "Right", "Decimal", "Bar" } },
                        leader = new { type = "string", description = "Tab leader: None, Dots, Dashes, Line, Heavy, MiddleDot (default: None)", @enum = new[] { "None", "Dots", "Dashes", "Line", "Heavy", "MiddleDot" } }
                    }
                }
            },
            applyToAllSections = new
            {
                type = "boolean",
                description = "Apply to all sections (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var headerText = arguments?["headerText"]?.GetValue<string>();
        var footerText = arguments?["footerText"]?.GetValue<string>();
        var headerLeft = arguments?["headerLeft"]?.GetValue<string>();
        var headerCenter = arguments?["headerCenter"]?.GetValue<string>();
        var headerRight = arguments?["headerRight"]?.GetValue<string>();
        var footerLeft = arguments?["footerLeft"]?.GetValue<string>();
        var footerCenter = arguments?["footerCenter"]?.GetValue<string>();
        var footerRight = arguments?["footerRight"]?.GetValue<string>();
        var headerAlignment = arguments?["headerAlignment"]?.GetValue<string>() ?? "center";
        var footerAlignment = arguments?["footerAlignment"]?.GetValue<string>() ?? "center";
        var headerFontName = arguments?["headerFontName"]?.GetValue<string>();
        var headerFontNameAscii = arguments?["headerFontNameAscii"]?.GetValue<string>();
        var headerFontNameFarEast = arguments?["headerFontNameFarEast"]?.GetValue<string>();
        var headerFontSize = arguments?["headerFontSize"]?.GetValue<double?>();
        var footerFontName = arguments?["footerFontName"]?.GetValue<string>();
        var footerFontNameAscii = arguments?["footerFontNameAscii"]?.GetValue<string>();
        var footerFontNameFarEast = arguments?["footerFontNameFarEast"]?.GetValue<string>();
        var footerFontSize = arguments?["footerFontSize"]?.GetValue<double?>();
        var includePageNumber = arguments?["includePageNumber"]?.GetValue<bool>() ?? false;
        var pageNumberFormat = arguments?["pageNumberFormat"]?.GetValue<string>() ?? "simple";
        var footerTemplate = arguments?["footerTemplate"]?.GetValue<string>();
        var headerBorder = arguments?["headerBorder"]?.GetValue<bool>() ?? true;
        var footerBorder = arguments?["footerBorder"]?.GetValue<bool>() ?? true;
        var headerLineStyle = arguments?["headerLineStyle"]?.GetValue<string>() ?? "border";
        var headerLineWidth = arguments?["headerLineWidth"]?.GetValue<double?>();
        var headerLineColor = arguments?["headerLineColor"]?.GetValue<string>() ?? "000000";
        var headerTabStops = arguments?["headerTabStops"]?.AsArray();
        var headerImagePath = arguments?["headerImagePath"]?.GetValue<string>();
        var headerImageWidth = arguments?["headerImageWidth"]?.GetValue<double?>() ?? 50;
        var headerImageHeight = arguments?["headerImageHeight"]?.GetValue<double?>();
        var headerImageAlignment = arguments?["headerImageAlignment"]?.GetValue<string>() ?? "left";
        var footerImagePath = arguments?["footerImagePath"]?.GetValue<string>();
        var footerImageWidth = arguments?["footerImageWidth"]?.GetValue<double?>() ?? 20;
        var footerImageHeight = arguments?["footerImageHeight"]?.GetValue<double?>();
        var footerImageAlignment = arguments?["footerImageAlignment"]?.GetValue<string>() ?? "left";
        var footerLineStyle = arguments?["footerLineStyle"]?.GetValue<string>() ?? "border";
        var footerLineWidth = arguments?["footerLineWidth"]?.GetValue<double?>();
        var footerLineColor = arguments?["footerLineColor"]?.GetValue<string>() ?? "000000";
        var footerLinePosition = arguments?["footerLinePosition"]?.GetValue<string>() ?? "above";
        var footerTabStops = arguments?["footerTabStops"]?.AsArray();
        var applyToAllSections = arguments?["applyToAllSections"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);

        var sections = applyToAllSections ? doc.Sections.Cast<Section>() : new[] { doc.Sections[0] };

        foreach (Section section in sections)
        {
            // Set header
            bool hasHeader = !string.IsNullOrEmpty(headerText) || !string.IsNullOrEmpty(headerLeft) || 
                            !string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight);
            
            if (hasHeader)
            {
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }
                else
                {
                    header.RemoveAllChildren();
                }

                // Check if using three-part layout
                bool useThreePart = !string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight);
                
                if (useThreePart)
                {
                    // Three-part header on SAME LINE using tab stops
                    var headerPara = new Paragraph(doc);
                    
                    // Set up tab stops: use custom if provided, otherwise auto-calculate
                    if (headerTabStops != null && headerTabStops.Count > 0)
                    {
                        // Use custom tab stops
                        foreach (var tabStopJson in headerTabStops)
                        {
                            var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
                            var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "Left";
                            var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "None";
                            
                            var alignment = alignmentStr switch
                            {
                                "Center" => Aspose.Words.TabAlignment.Center,
                                "Right" => Aspose.Words.TabAlignment.Right,
                                "Decimal" => Aspose.Words.TabAlignment.Decimal,
                                "Bar" => Aspose.Words.TabAlignment.Bar,
                                _ => Aspose.Words.TabAlignment.Left
                            };
                            
                            var leader = leaderStr switch
                            {
                                "Dots" => Aspose.Words.TabLeader.Dots,
                                "Dashes" => Aspose.Words.TabLeader.Dashes,
                                "Line" => Aspose.Words.TabLeader.Line,
                                "Heavy" => Aspose.Words.TabLeader.Heavy,
                                "MiddleDot" => Aspose.Words.TabLeader.MiddleDot,
                                _ => Aspose.Words.TabLeader.None
                            };
                            
                            headerPara.ParagraphFormat.TabStops.Add(new Aspose.Words.TabStop(position, alignment, leader));
                        }
                    }
                    else
                    {
                        // Auto-calculate tab stops
                        // Tab stops in Aspose.Words are measured from the LEFT PAGE EDGE (not left margin)
                        var pageWidth = section.PageSetup.PageWidth;
                        var rightMargin = section.PageSetup.RightMargin;
                        
                        // Only add center tab if there's center content
                        if (!string.IsNullOrEmpty(headerCenter))
                        {
                            var centerPos = pageWidth / 2;
                            headerPara.ParagraphFormat.TabStops.Add(new Aspose.Words.TabStop(centerPos, Aspose.Words.TabAlignment.Center, Aspose.Words.TabLeader.None));
                        }
                        
                        // Always add right tab if there's right content
                        if (!string.IsNullOrEmpty(headerRight))
                        {
                            var rightPos = pageWidth - rightMargin;
                            headerPara.ParagraphFormat.TabStops.Add(new Aspose.Words.TabStop(rightPos, Aspose.Words.TabAlignment.Right, Aspose.Words.TabLeader.None));
                        }
                    }
                    
                    // Add left text
                    if (!string.IsNullOrEmpty(headerLeft))
                    {
                        var leftRun = new Run(doc, headerLeft);
                        ApplyFontSettings(leftRun, headerFontName, headerFontNameAscii, headerFontNameFarEast, headerFontSize);
                        headerPara.AppendChild(leftRun);
                    }
                    
                    // Add center text (with tab before it)
                    if (!string.IsNullOrEmpty(headerCenter))
                    {
                        headerPara.AppendChild(new Run(doc, "\t"));
                        var centerRun = new Run(doc, headerCenter);
                        ApplyFontSettings(centerRun, headerFontName, headerFontNameAscii, headerFontNameFarEast, headerFontSize);
                        headerPara.AppendChild(centerRun);
                    }
                    
                    // Add right text (with tab before it)
                    if (!string.IsNullOrEmpty(headerRight))
                    {
                        headerPara.AppendChild(new Run(doc, "\t"));
                        var rightRun = new Run(doc, headerRight);
                        ApplyFontSettings(rightRun, headerFontName, headerFontNameAscii, headerFontNameFarEast, headerFontSize);
                        headerPara.AppendChild(rightRun);
                    }
                    
                    header.AppendChild(headerPara);
                }
                else if (!string.IsNullOrEmpty(headerText))
                {
                    // Simple single header text (backward compatible)
                    var headerPara = new Paragraph(doc);
                    headerPara.ParagraphFormat.Alignment = GetAlignment(headerAlignment);
                    var headerRun = new Run(doc, headerText);
                    ApplyFontSettings(headerRun, headerFontName, headerFontNameAscii, headerFontNameFarEast, headerFontSize);
                    headerPara.AppendChild(headerRun);
                    header.AppendChild(headerPara);
                }
                
                // Add header image if specified
                if (!string.IsNullOrEmpty(headerImagePath) && header != null)
                {
                    var imagePara = new Paragraph(doc);
                    imagePara.ParagraphFormat.Alignment = GetAlignment(headerImageAlignment);
                    
                    // First append paragraph to header, then insert image
                    header.AppendChild(imagePara);
                    
                    var builder = new DocumentBuilder(doc);
                    builder.MoveTo(imagePara);
                    
                    var shape = builder.InsertImage(headerImagePath);
                    shape.Width = headerImageWidth;
                    if (headerImageHeight.HasValue)
                        shape.Height = headerImageHeight.Value;
                    else
                        shape.AspectRatioLocked = true; // Maintain aspect ratio
                }
                
                // Set header line (border or shape)
                if (header != null && headerBorder)
                {
                    if (headerLineStyle == "shape")
                    {
                        // Add a graphic line using Shape
                        var linePara = new Paragraph(doc);
                        
                        // CRITICAL: Set paragraph spacing to 0 to avoid blank lines
                        linePara.ParagraphFormat.SpaceBefore = 0;
                        linePara.ParagraphFormat.SpaceAfter = 0;
                        linePara.ParagraphFormat.LineSpacing = 1; // Minimum line spacing
                        linePara.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
                        
                        var contentWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
                        var lineWidth = headerLineWidth ?? 1.0;
                        
                        var shape = new Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.Line);
                        shape.Width = contentWidth;
                        shape.Height = 0; // Horizontal line
                        shape.StrokeWeight = lineWidth;
                        shape.StrokeColor = ParseColor(headerLineColor);
                        
                        // Set shape to inline to avoid taking extra space
                        shape.WrapType = Aspose.Words.Drawing.WrapType.Inline;
                        
                        linePara.AppendChild(shape);
                        header.AppendChild(linePara);
                    }
                    else
                    {
                        // Use paragraph border (default, backward compatible)
                        var firstPara = header.FirstParagraph;
                        if (firstPara != null)
                        {
                            var lineWidth = headerLineWidth ?? 0.5;
                            firstPara.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;
                            firstPara.ParagraphFormat.Borders.Bottom.LineWidth = lineWidth;
                            firstPara.ParagraphFormat.Borders.Bottom.Color = ParseColor(headerLineColor);
                        }
                    }
                }
                else if (header != null && !headerBorder)
                {
                    // Remove border if explicitly disabled
                    var firstPara = header.FirstParagraph;
                    if (firstPara != null)
                    {
                        firstPara.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.None;
                    }
                }
            }

            // Set footer
            bool hasFooter = !string.IsNullOrEmpty(footerText) || !string.IsNullOrEmpty(footerLeft) || 
                            !string.IsNullOrEmpty(footerCenter) || !string.IsNullOrEmpty(footerRight) || includePageNumber;
            
            if (hasFooter)
            {
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                    section.HeadersFooters.Add(footer);
                }
                else
                {
                    footer.RemoveAllChildren();
                }

                // Check if using three-part layout
                bool useThreePart = !string.IsNullOrEmpty(footerLeft) || !string.IsNullOrEmpty(footerCenter) || !string.IsNullOrEmpty(footerRight);
                
                if (useThreePart)
                {
                    // Three-part footer on SAME LINE using tab stops
                    var footerPara = new Paragraph(doc);
                    
                    // Set up tab stops: use custom if provided, otherwise auto-calculate
                    if (footerTabStops != null && footerTabStops.Count > 0)
                    {
                        // Use custom tab stops
                        foreach (var tabStopJson in footerTabStops)
                        {
                            var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
                            var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "Left";
                            var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "None";
                            
                            var alignment = alignmentStr switch
                            {
                                "Center" => Aspose.Words.TabAlignment.Center,
                                "Right" => Aspose.Words.TabAlignment.Right,
                                "Decimal" => Aspose.Words.TabAlignment.Decimal,
                                "Bar" => Aspose.Words.TabAlignment.Bar,
                                _ => Aspose.Words.TabAlignment.Left
                            };
                            
                            var leader = leaderStr switch
                            {
                                "Dots" => Aspose.Words.TabLeader.Dots,
                                "Dashes" => Aspose.Words.TabLeader.Dashes,
                                "Line" => Aspose.Words.TabLeader.Line,
                                "Heavy" => Aspose.Words.TabLeader.Heavy,
                                "MiddleDot" => Aspose.Words.TabLeader.MiddleDot,
                                _ => Aspose.Words.TabLeader.None
                            };
                            
                            footerPara.ParagraphFormat.TabStops.Add(new Aspose.Words.TabStop(position, alignment, leader));
                        }
                    }
                    else if (footerTabStops == null)
                    {
                        // Auto-calculate tab stops (only if footerTabStops is not explicitly set)
                        // Tab stops in Aspose.Words are measured from the LEFT PAGE EDGE (not left margin)
                        var pageWidth = section.PageSetup.PageWidth;
                        var rightMargin = section.PageSetup.RightMargin;
                        
                        // Only add center tab if there's center content
                        if (!string.IsNullOrEmpty(footerCenter))
                        {
                            var centerPos = pageWidth / 2;
                            footerPara.ParagraphFormat.TabStops.Add(new Aspose.Words.TabStop(centerPos, Aspose.Words.TabAlignment.Center, Aspose.Words.TabLeader.None));
                        }
                        
                        // Always add right tab if there's right content
                        if (!string.IsNullOrEmpty(footerRight))
                        {
                            var rightPos = pageWidth - rightMargin;
                            footerPara.ParagraphFormat.TabStops.Add(new Aspose.Words.TabStop(rightPos, Aspose.Words.TabAlignment.Right, Aspose.Words.TabLeader.None));
                        }
                    }
                    // If footerTabStops is an empty array, don't add any tab stops
                    
                    // Add left text/content
                    if (!string.IsNullOrEmpty(footerLeft))
                    {
                        // Check if footerLeft contains placeholders
                        if (footerLeft.Contains("{PAGE}") || footerLeft.Contains("{NUMPAGES}") || footerLeft.Contains("{SECTION}"))
                        {
                            ProcessFooterTemplate(doc, footerPara, footerLeft, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                        }
                        else
                        {
                            var leftRun = new Run(doc, footerLeft);
                            ApplyFontSettings(leftRun, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                            footerPara.AppendChild(leftRun);
                        }
                    }
                    
                    // Add center text/content (with tab before it)
                    if (!string.IsNullOrEmpty(footerCenter) || includePageNumber)
                    {
                        footerPara.AppendChild(new Run(doc, "\t"));
                        
                        if (!string.IsNullOrEmpty(footerCenter))
                        {
                            // Check if footerCenter contains placeholders
                            if (footerCenter.Contains("{PAGE}") || footerCenter.Contains("{NUMPAGES}") || footerCenter.Contains("{SECTION}"))
                            {
                                ProcessFooterTemplate(doc, footerPara, footerCenter, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                            }
                            else
                            {
                                var centerRun = new Run(doc, footerCenter);
                                ApplyFontSettings(centerRun, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                                footerPara.AppendChild(centerRun);
                            }
                        }
                        
                        if (includePageNumber)
                        {
                            AddPageNumber(doc, footerPara, pageNumberFormat, footerTemplate, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                        }
                    }
                    
                    // Add right text/content (with tab before it)
                    if (!string.IsNullOrEmpty(footerRight))
                    {
                        footerPara.AppendChild(new Run(doc, "\t"));
                        
                        // Check if footerRight contains placeholders
                        if (footerRight.Contains("{PAGE}") || footerRight.Contains("{NUMPAGES}") || footerRight.Contains("{SECTION}"))
                        {
                            ProcessFooterTemplate(doc, footerPara, footerRight, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                        }
                        else
                        {
                            var rightRun = new Run(doc, footerRight);
                            ApplyFontSettings(rightRun, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                            footerPara.AppendChild(rightRun);
                        }
                    }
                    
                    footer.AppendChild(footerPara);
                }
                else
                {
                    // Simple single footer (backward compatible)
                    var footerPara = new Paragraph(doc);
                    footerPara.ParagraphFormat.Alignment = GetAlignment(footerAlignment);

                    if (!string.IsNullOrEmpty(footerText))
                    {
                        var footerRun = new Run(doc, footerText);
                        ApplyFontSettings(footerRun, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                        footerPara.AppendChild(footerRun);
                        
                        if (includePageNumber)
                        {
                            footerPara.AppendChild(new Run(doc, " "));
                        }
                    }

                    if (includePageNumber)
                    {
                        AddPageNumber(doc, footerPara, pageNumberFormat, footerTemplate, footerFontName, footerFontNameAscii, footerFontNameFarEast, footerFontSize);
                    }

                    footer.AppendChild(footerPara);
                }
                
                // Add footer image if specified
                if (!string.IsNullOrEmpty(footerImagePath) && footer != null)
                {
                    // For footer with three-part layout, insert image at the beginning of the first paragraph
                    bool hasThreePartFooter = !string.IsNullOrEmpty(footerLeft) || !string.IsNullOrEmpty(footerCenter) || !string.IsNullOrEmpty(footerRight);
                    
                    if (hasThreePartFooter && footer.FirstParagraph != null)
                    {
                        // Insert image inline at the beginning of the footer paragraph
                        var builder = new DocumentBuilder(doc);
                        builder.MoveTo(footer.FirstParagraph);
                        builder.MoveToDocumentStart();
                        
                        var shape = builder.InsertImage(footerImagePath);
                        shape.Width = footerImageWidth;
                        if (footerImageHeight.HasValue)
                            shape.Height = footerImageHeight.Value;
                        else
                            shape.AspectRatioLocked = true; // Maintain aspect ratio
                        
                        // Add a space after the image
                        footer.FirstParagraph.InsertBefore(new Run(doc, " "), footer.FirstParagraph.FirstChild);
                        footer.FirstParagraph.InsertBefore(shape, footer.FirstParagraph.FirstChild);
                    }
                    else
                    {
                        // Add image in a separate paragraph
                        var imagePara = new Paragraph(doc);
                        imagePara.ParagraphFormat.Alignment = GetAlignment(footerImageAlignment);
                        
                        // First prepend paragraph to footer, then insert image
                        footer.PrependChild(imagePara);
                        
                        var builder = new DocumentBuilder(doc);
                        builder.MoveTo(imagePara);
                        
                        var shape = builder.InsertImage(footerImagePath);
                        shape.Width = footerImageWidth;
                        if (footerImageHeight.HasValue)
                            shape.Height = footerImageHeight.Value;
                        else
                            shape.AspectRatioLocked = true; // Maintain aspect ratio
                    }
                }
                
                // Set footer line (border or shape)
                if (footer != null && footerBorder)
                {
                    if (footerLineStyle == "shape")
                    {
                        // Add a graphic line using Shape
                        var linePara = new Paragraph(doc);
                        
                        // CRITICAL: Set paragraph spacing to 0 to avoid blank lines
                        linePara.ParagraphFormat.SpaceBefore = 0;
                        linePara.ParagraphFormat.SpaceAfter = 0;
                        linePara.ParagraphFormat.LineSpacing = 1; // Minimum line spacing
                        linePara.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
                        
                        var contentWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
                        var lineWidth = footerLineWidth ?? 1.0;
                        
                        var shape = new Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.Line);
                        shape.Width = contentWidth;
                        shape.Height = 0; // Horizontal line
                        shape.StrokeWeight = lineWidth;
                        shape.StrokeColor = ParseColor(footerLineColor);
                        
                        // Set shape to inline to avoid taking extra space
                        shape.WrapType = Aspose.Words.Drawing.WrapType.Inline;
                        
                        linePara.AppendChild(shape);
                        
                        // Insert line based on position
                        if (footerLinePosition == "above")
                        {
                            // Insert line at the beginning (above content)
                            footer.PrependChild(linePara);
                        }
                        else
                        {
                            // Insert line at the end (below content)
                            footer.AppendChild(linePara);
                        }
                    }
                    else
                    {
                        // Use paragraph border (default, backward compatible)
                        var firstPara = footer.FirstParagraph;
                        if (firstPara != null)
                        {
                            var lineWidth = footerLineWidth ?? 0.5;
                            firstPara.ParagraphFormat.Borders.Top.LineStyle = LineStyle.Single;
                            firstPara.ParagraphFormat.Borders.Top.LineWidth = lineWidth;
                            firstPara.ParagraphFormat.Borders.Top.Color = ParseColor(footerLineColor);
                        }
                    }
                }
                else if (footer != null && !footerBorder)
                {
                    // Remove border if explicitly disabled
                    var firstPara = footer.FirstParagraph;
                    if (firstPara != null)
                    {
                        firstPara.ParagraphFormat.Borders.Top.LineStyle = LineStyle.None;
                    }
                }
            }
        }

        doc.Save(outputPath);

        var sectionsUpdated = applyToAllSections ? doc.Sections.Count : 1;
        var result = $"成功設定 {sectionsUpdated} 個節的頁首頁尾\n";
        
        // Header info
        bool hasHeaderInfo = !string.IsNullOrEmpty(headerText) || !string.IsNullOrEmpty(headerLeft) || 
                             !string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight);
        if (hasHeaderInfo)
        {
            result += "頁首: ";
            if (!string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight))
            {
                result += "[三段式] ";
                if (!string.IsNullOrEmpty(headerLeft)) result += $"左={headerLeft} ";
                if (!string.IsNullOrEmpty(headerCenter)) result += $"中={headerCenter} ";
                if (!string.IsNullOrEmpty(headerRight)) result += $"右={headerRight} ";
            }
            else if (!string.IsNullOrEmpty(headerText))
            {
                result += $"{headerText} ({headerAlignment}) ";
            }
            
            if (!string.IsNullOrEmpty(headerFontName)) result += $"字型={headerFontName} ";
            if (headerFontSize.HasValue) result += $"字號={headerFontSize.Value}pt ";
            result += "\n";
        }
        
        // Footer info
        bool hasFooterInfo = !string.IsNullOrEmpty(footerText) || !string.IsNullOrEmpty(footerLeft) || 
                             !string.IsNullOrEmpty(footerCenter) || !string.IsNullOrEmpty(footerRight) || includePageNumber;
        if (hasFooterInfo)
        {
            result += "頁尾: ";
            if (!string.IsNullOrEmpty(footerLeft) || !string.IsNullOrEmpty(footerCenter) || !string.IsNullOrEmpty(footerRight))
            {
                result += "[三段式] ";
                if (!string.IsNullOrEmpty(footerLeft)) result += $"左={footerLeft} ";
                if (!string.IsNullOrEmpty(footerCenter)) result += $"中={footerCenter} ";
                if (!string.IsNullOrEmpty(footerRight)) result += $"右={footerRight} ";
            }
            else if (!string.IsNullOrEmpty(footerText))
            {
                result += $"{footerText} ({footerAlignment}) ";
            }
            
            if (includePageNumber)
            {
                if (pageNumberFormat == "custom" && !string.IsNullOrEmpty(footerTemplate))
                    result += $"[自訂格式: {footerTemplate}] ";
                else
                    result += $"[頁碼格式: {pageNumberFormat}] ";
            }
            
            if (!string.IsNullOrEmpty(footerFontName)) result += $"字型={footerFontName} ";
            if (footerFontSize.HasValue) result += $"字號={footerFontSize.Value}pt ";
            result += "\n";
        }
        
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private void ApplyFontSettings(Run run, string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize)
    {
        // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
        if (!string.IsNullOrEmpty(fontNameAscii))
            run.Font.NameAscii = fontNameAscii;
        
        if (!string.IsNullOrEmpty(fontNameFarEast))
            run.Font.NameFarEast = fontNameFarEast;
        
        if (!string.IsNullOrEmpty(fontName))
        {
            // If fontNameAscii/FarEast are not set, use fontName for both
            if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
            {
                run.Font.Name = fontName;
            }
            else
            {
                // If only one is set, use fontName as fallback for the other
                if (string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontName;
                if (string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontName;
            }
        }
        
        if (fontSize.HasValue)
        {
            run.Font.Size = fontSize.Value;
        }
    }

    private void AddPageNumber(Document doc, Paragraph para, string format, string? template, string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize)
    {
        switch (format)
        {
            case "custom":
                if (!string.IsNullOrEmpty(template))
                {
                    ProcessFooterTemplate(doc, para, template, fontName, fontNameAscii, fontNameFarEast, fontSize);
                }
                else
                {
                    var run = new Run(doc, "Page ");
                    ApplyFontSettings(run, fontName, fontNameAscii, fontNameFarEast, fontSize);
                    para.AppendChild(run);
                    para.AppendField("PAGE", "");
                }
                break;
            case "simple":
                var simpleRun = new Run(doc, "Page ");
                ApplyFontSettings(simpleRun, fontName, fontNameAscii, fontNameFarEast, fontSize);
                para.AppendChild(simpleRun);
                para.AppendField("PAGE", "");
                break;
            case "total":
                var totalRun1 = new Run(doc, "Page ");
                ApplyFontSettings(totalRun1, fontName, fontNameAscii, fontNameFarEast, fontSize);
                para.AppendChild(totalRun1);
                para.AppendField("PAGE", "");
                var totalRun2 = new Run(doc, " of ");
                ApplyFontSettings(totalRun2, fontName, fontNameAscii, fontNameFarEast, fontSize);
                para.AppendChild(totalRun2);
                para.AppendField("NUMPAGES", "");
                break;
            case "chinese":
                var chineseRun1 = new Run(doc, "第 ");
                ApplyFontSettings(chineseRun1, fontName, fontNameAscii, fontNameFarEast, fontSize);
                para.AppendChild(chineseRun1);
                para.AppendField("PAGE", "");
                var chineseRun2 = new Run(doc, " 頁，共 ");
                ApplyFontSettings(chineseRun2, fontName, fontNameAscii, fontNameFarEast, fontSize);
                para.AppendChild(chineseRun2);
                para.AppendField("NUMPAGES", "");
                var chineseRun3 = new Run(doc, " 頁");
                ApplyFontSettings(chineseRun3, fontName, fontNameAscii, fontNameFarEast, fontSize);
                para.AppendChild(chineseRun3);
                break;
        }
    }

    private void ProcessFooterTemplate(Document doc, Paragraph para, string template, string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize)
    {
        // Split template by placeholders and process each part
        var parts = System.Text.RegularExpressions.Regex.Split(template, @"(\{PAGE\}|\{NUMPAGES\}|\{SECTION\})");
        
        foreach (var part in parts)
        {
            if (string.IsNullOrEmpty(part))
                continue;
                
            switch (part)
            {
                case "{PAGE}":
                    para.AppendField("PAGE", "");
                    break;
                case "{NUMPAGES}":
                    para.AppendField("NUMPAGES", "");
                    break;
                case "{SECTION}":
                    para.AppendField("SECTION", "");
                    break;
                default:
                    // Regular text
                    var run = new Run(doc, part);
                    ApplyFontSettings(run, fontName, fontNameAscii, fontNameFarEast, fontSize);
                    para.AppendChild(run);
                    break;
            }
        }
    }

    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "right" => ParagraphAlignment.Right,
            "center" => ParagraphAlignment.Center,
            _ => ParagraphAlignment.Center
        };
    }

    private System.Drawing.Color ParseColor(string hexColor)
    {
        // Remove '#' if present
        hexColor = hexColor.TrimStart('#');
        
        // Ensure it's 6 characters (RGB)
        if (hexColor.Length != 6)
        {
            return System.Drawing.Color.Black; // Default to black if invalid
        }
        
        try
        {
            int r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
            int g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
            int b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        catch
        {
            return System.Drawing.Color.Black; // Default to black if parsing fails
        }
    }
}

