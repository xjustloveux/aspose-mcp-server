using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Settings;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for Word file operations (create, create_from_template, convert, merge, split)
/// </summary>
public class WordFileTool : IAsposeTool
{
    public string Description =>
        @"Perform file operations on Word documents. Supports 5 operations: create, create_from_template, convert, merge, split.

Usage examples:
- Create document: word_file(operation='create', outputPath='new.docx')
- Create from template: word_file(operation='create_from_template', templatePath='template.docx', outputPath='output.docx', data={'Name':'John','Items':[{'Product':'Apple','Price':25}]})
- Convert format: word_file(operation='convert', path='doc.docx', outputPath='doc.pdf', format='pdf')
- Merge documents: word_file(operation='merge', inputPaths=['doc1.docx','doc2.docx'], outputPath='merged.docx')
- Split document: word_file(operation='split', path='doc.docx', outputDir='output/', splitBy='page')

Template syntax (LINQ Reporting Engine, use 'ds' as data source prefix):
- Simple value: <<[ds.Name]>>
- Nested object: <<[ds.Customer.Address.City]>>
- Array iteration: <<foreach [item in ds.Items]>><<[item.Product]>>: <<[item.Price]>><</foreach>>
- Root array iteration: <<foreach [item in ds]>><<[item.Name]>><</foreach>>
- Conditional: <<if [ds.Total > 1000]>>VIP<<else>>Normal<</if>>
- Image: <<image [ds.ImageBytes]>>";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'create': Create a new document (required params: outputPath)
- 'create_from_template': Create from template (required params: templatePath, outputPath)
- 'convert': Convert document format (required params: path, outputPath, format)
- 'merge': Merge multiple documents (required params: inputPaths, outputPath)
- 'split': Split document (required params: path, outputDir)",
                @enum = new[] { "create", "create_from_template", "convert", "merge", "split" }
            },
            path = new
            {
                type = "string",
                description = "Input file path (required for convert and split operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (required for create, create_from_template, convert, and merge operations)"
            },
            templatePath = new
            {
                type = "string",
                description = "Template file path (required for create_from_template)"
            },
            data = new
            {
                type = "object",
                description = @"Data object for template rendering using LINQ Reporting Engine.
Use 'ds' prefix to access data: <<[ds.PropertyName]>>, <<foreach [item in ds.Items]>>...<</foreach>>.
For root arrays: <<foreach [item in ds]>><<[item.Name]>><</foreach>>"
            },
            format = new
            {
                type = "string",
                description = "Output format: pdf, html, docx, txt, rtf, odt, epub, xps (for convert)"
            },
            inputPaths = new
            {
                type = "array",
                description = "Array of input file paths to merge (for merge)",
                items = new { type = "string" }
            },
            importFormatMode = new
            {
                type = "string",
                description =
                    "Format mode when merging: KeepSourceFormatting (default), UseDestinationStyles, KeepDifferentStyles (for merge)",
                @enum = new[] { "KeepSourceFormatting", "UseDestinationStyles", "KeepDifferentStyles" }
            },
            unlinkHeadersFooters = new
            {
                type = "boolean",
                description = "Unlink headers/footers after merge to prevent confusion (for merge, default: false)"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory for split files (for split)"
            },
            splitBy = new
            {
                type = "string",
                description = "Split by: section, page (default: section)",
                @enum = new[] { "section", "page" }
            },
            content = new
            {
                type = "string",
                description = "Initial content (for create, optional)"
            },
            skipInitialContent = new
            {
                type = "boolean",
                description = "Create blank document (for create, default: false)"
            },
            marginTop = new
            {
                type = "number",
                description = "Top margin in points (for create, default: 70.87)"
            },
            marginBottom = new
            {
                type = "number",
                description = "Bottom margin in points (for create, default: 70.87)"
            },
            marginLeft = new
            {
                type = "number",
                description = "Left margin in points (for create, default: 70.87)"
            },
            marginRight = new
            {
                type = "number",
                description = "Right margin in points (for create, default: 70.87)"
            },
            compatibilityMode = new
            {
                type = "string",
                description = "Word compatibility mode: Word2019, Word2016, Word2013, Word2010, Word2007 (for create)",
                @enum = new[] { "Word2019", "Word2016", "Word2013", "Word2010", "Word2007" }
            },
            paperSize = new
            {
                type = "string",
                description =
                    "Predefined paper size: A4, Letter, A3, Legal (for create, default: A4). Overrides pageWidth/pageHeight if specified.",
                @enum = new[] { "A4", "Letter", "A3", "Legal" }
            },
            pageWidth = new
            {
                type = "number",
                description =
                    "Page width in points (for create, e.g., 595.3 for A4, 612 for Letter). Default: 595.3 (A4)"
            },
            pageHeight = new
            {
                type = "number",
                description =
                    "Page height in points (for create, e.g., 841.9 for A4, 792 for Letter). Default: 841.9 (A4)"
            },
            headerDistance = new
            {
                type = "number",
                description =
                    "Header distance from page top in points (for create, e.g., 45.35 pt = 1.6 cm). Default: 35.4"
            },
            footerDistance = new
            {
                type = "number",
                description =
                    "Footer distance from page bottom in points (for create, e.g., 45.35 pt = 1.6 cm). Default: 35.4"
            }
        },
        required = new[] { "operation" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "create" => await CreateDocument(arguments),
            "create_from_template" => await CreateFromTemplate(arguments),
            "convert" => await ConvertDocument(arguments),
            "merge" => await MergeDocuments(arguments),
            "split" => await SplitDocument(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new Word document
    /// </summary>
    /// <param name="arguments">JSON arguments containing outputPath, optional content</param>
    /// <returns>Success message with file path</returns>
    private Task<string> CreateDocument(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);
            var content = ArgumentHelper.GetStringNullable(arguments, "content");
            var skipInitialContent = ArgumentHelper.GetBool(arguments, "skipInitialContent", false);
            var marginTop = ArgumentHelper.GetDouble(arguments, "marginTop", 70.87);
            var marginBottom = ArgumentHelper.GetDouble(arguments, "marginBottom", 70.87);
            var marginLeft = ArgumentHelper.GetDouble(arguments, "marginLeft", 70.87);
            var marginRight = ArgumentHelper.GetDouble(arguments, "marginRight", 70.87);
            var compatibilityMode = ArgumentHelper.GetString(arguments, "compatibilityMode", "Word2019");
            var pageWidth = ArgumentHelper.GetDoubleNullable(arguments, "pageWidth");
            var pageHeight = ArgumentHelper.GetDoubleNullable(arguments, "pageHeight");
            var paperSize = ArgumentHelper.GetString(arguments, "paperSize", "A4");
            var headerDistance = ArgumentHelper.GetDouble(arguments, "headerDistance", 35.4);
            var footerDistance = ArgumentHelper.GetDouble(arguments, "footerDistance", 35.4);

            var doc = new Document();

            var wordVersion = compatibilityMode switch
            {
                "Word2019" => MsWordVersion.Word2019,
                "Word2016" => MsWordVersion.Word2016,
                "Word2013" => MsWordVersion.Word2013,
                "Word2010" => MsWordVersion.Word2010,
                "Word2007" => MsWordVersion.Word2007,
                _ => MsWordVersion.Word2019
            };
            doc.CompatibilityOptions.OptimizeFor(wordVersion);

            var section = doc.FirstSection;
            if (section != null)
            {
                var pageSetup = section.PageSetup;

                // Set page size (paper size or custom dimensions)
                if (!string.IsNullOrEmpty(paperSize) && pageWidth == null && pageHeight == null)
                {
                    // Use Aspose's built-in PaperSize enum for accurate dimensions
                    pageSetup.PaperSize = paperSize.ToUpper() switch
                    {
                        "A4" => PaperSize.A4,
                        "LETTER" => PaperSize.Letter,
                        "A3" => PaperSize.A3,
                        "LEGAL" => PaperSize.Legal,
                        _ => PaperSize.A4
                    };
                }
                else if (pageWidth != null || pageHeight != null)
                {
                    // Use custom dimensions if specified
                    pageSetup.PaperSize = PaperSize.Custom;
                    pageSetup.PageWidth = pageWidth ?? 595.3; // Default to A4 width
                    pageSetup.PageHeight = pageHeight ?? 841.9; // Default to A4 height
                }
                else
                {
                    // Default to A4
                    pageSetup.PaperSize = PaperSize.A4;
                }

                // Set margins
                pageSetup.TopMargin = marginTop;
                pageSetup.BottomMargin = marginBottom;
                pageSetup.LeftMargin = marginLeft;
                pageSetup.RightMargin = marginRight;

                // Set header/footer distance
                pageSetup.HeaderDistance = headerDistance;
                pageSetup.FooterDistance = footerDistance;
            }

            var builder = new DocumentBuilder(doc);

            if (skipInitialContent)
            {
                if (doc.FirstSection is { Body: not null })
                {
                    doc.FirstSection.Body.RemoveAllChildren();
                    var firstPara = new Paragraph(doc)
                    {
                        ParagraphFormat =
                        {
                            SpaceBefore = 0,
                            SpaceAfter = 0,
                            LineSpacing = 12
                        }
                    };
                    doc.FirstSection.Body.AppendChild(firstPara);
                }
            }
            else if (!string.IsNullOrEmpty(content))
            {
                builder.Write(content);
            }

            doc.Save(outputPath);
            return $"Word document created successfully at: {outputPath}";
        });
    }

    /// <summary>
    ///     Creates a document from a template using LINQ Reporting Engine
    /// </summary>
    /// <param name="arguments">JSON arguments containing templatePath, outputPath, data object</param>
    /// <returns>Success message with output path</returns>
    private Task<string> CreateFromTemplate(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var templatePath = ArgumentHelper.GetString(arguments, "templatePath");
            var outputPath = ArgumentHelper.GetString(arguments, "outputPath");

            SecurityHelper.ValidateFilePath(templatePath, "templatePath", true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Get data object
            var dataNode = arguments?["data"];
            if (dataNode == null)
                throw new ArgumentException("data parameter is required for create_from_template");

            var doc = new Document(templatePath);
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Use Aspose's JsonDataSource for proper JSON handling
            var jsonString = dataNode.ToJsonString();
            using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
            var loadOptions = new JsonDataLoadOptions
            {
                ExactDateTimeParseFormats = new List<string> { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss" },
                SimpleValueParseMode = JsonSimpleValueParseMode.Strict
            };
            var dataSource = new JsonDataSource(jsonStream, loadOptions);

            // Build report with the data
            engine.BuildReport(doc, dataSource, "ds");

            doc.Save(outputPath);
            return $"Document created from template using LINQ Reporting Engine: {outputPath}";
        });
    }

    /// <summary>
    ///     Converts document to another format
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, outputPath, format</param>
    /// <returns>Success message with output path</returns>
    private Task<string> ConvertDocument(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidatePath(arguments, "outputPath");
            var formatParam = ArgumentHelper.GetStringNullable(arguments, "format");

            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            // Auto-infer format from file extension if not specified
            var format = formatParam?.ToLower();
            if (string.IsNullOrEmpty(format))
            {
                var extension = Path.GetExtension(outputPath).TrimStart('.').ToLower();
                format = extension switch
                {
                    "pdf" => "pdf",
                    "html" or "htm" => "html",
                    "docx" => "docx",
                    "doc" => "doc",
                    "txt" => "txt",
                    "rtf" => "rtf",
                    "odt" => "odt",
                    "epub" => "epub",
                    "xps" => "xps",
                    _ => throw new ArgumentException(
                        $"Cannot infer format from extension '.{extension}'. Please specify format parameter.")
                };
            }

            var doc = new Document(path);

            var saveFormat = format switch
            {
                "pdf" => SaveFormat.Pdf,
                "html" => SaveFormat.Html,
                "docx" => SaveFormat.Docx,
                "doc" => SaveFormat.Doc,
                "txt" => SaveFormat.Text,
                "rtf" => SaveFormat.Rtf,
                "odt" => SaveFormat.Odt,
                "epub" => SaveFormat.Epub,
                "xps" => SaveFormat.Xps,
                _ => throw new ArgumentException($"Unsupported format: {format}")
            };

            doc.Save(outputPath, saveFormat);
            return $"Document converted from {path} to {outputPath} ({format})";
        });
    }

    /// <summary>
    ///     Merges multiple documents into one
    /// </summary>
    /// <param name="arguments">JSON arguments containing sourcePaths array, outputPath</param>
    /// <returns>Success message with merged file path</returns>
    private Task<string> MergeDocuments(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var inputPathsArray = ArgumentHelper.GetArray(arguments, "inputPaths");
            var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
            var importFormatModeStr = ArgumentHelper.GetString(arguments, "importFormatMode", "KeepSourceFormatting");
            var unlinkHeadersFooters = ArgumentHelper.GetBool(arguments, "unlinkHeadersFooters", false);

            SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => p != null).ToList();

            if (inputPaths.Count == 0)
                throw new ArgumentException("At least one input path is required");

            foreach (var inputPath in inputPaths) SecurityHelper.ValidateFilePath(inputPath!, "inputPaths", true);

            // Parse import format mode
            var importFormatMode = importFormatModeStr switch
            {
                "UseDestinationStyles" => ImportFormatMode.UseDestinationStyles,
                "KeepDifferentStyles" => ImportFormatMode.KeepDifferentStyles,
                _ => ImportFormatMode.KeepSourceFormatting
            };

            var mergedDoc = new Document(inputPaths[0]);

            for (var i = 1; i < inputPaths.Count; i++)
            {
                var doc = new Document(inputPaths[i]);
                mergedDoc.AppendDocument(doc, importFormatMode);
            }

            // Unlink headers/footers if requested to prevent confusion after merge
            if (unlinkHeadersFooters)
                foreach (var section in mergedDoc.Sections.Cast<Section>())
                    section.HeadersFooters.LinkToPrevious(false);

            mergedDoc.Save(outputPath);
            return $"Merged {inputPaths.Count} documents into: {outputPath} (format mode: {importFormatModeStr})";
        });
    }

    /// <summary>
    ///     Splits document into multiple files
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, outputPath, splitBy (page or section)</param>
    /// <returns>Success message with split file count</returns>
    private Task<string> SplitDocument(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputDir = ArgumentHelper.GetString(arguments, "outputDir");
            var splitBy = ArgumentHelper.GetString(arguments, "splitBy", "section");

            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

            Directory.CreateDirectory(outputDir);

            var doc = new Document(path);
            var fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(path));

            if (splitBy.ToLower() == "section")
            {
                for (var i = 0; i < doc.Sections.Count; i++)
                {
                    // Use RemoveAllChildren + ImportNode for cleaner section cloning
                    var sectionDoc = new Document();
                    sectionDoc.RemoveAllChildren();
                    sectionDoc.AppendChild(sectionDoc.ImportNode(doc.Sections[i], true));

                    var outputPath = Path.Combine(outputDir, $"{fileBaseName}_section_{i + 1}.docx");
                    sectionDoc.Save(outputPath);
                }

                return $"Document split into {doc.Sections.Count} sections in: {outputDir}";
            }

            // For page split, update page layout first for accurate pagination
            doc.UpdatePageLayout();

            var pageCount = doc.PageCount;
            for (var i = 0; i < pageCount; i++)
            {
                var pageDoc = doc.ExtractPages(i, 1);
                var outputPath = Path.Combine(outputDir, $"{fileBaseName}_page_{i + 1}.docx");
                pageDoc.Save(outputPath);
            }

            return $"Document split into {pageCount} pages in: {outputDir}";
        });
    }
}