using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Replacing;
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
- Create from template: word_file(operation='create_from_template', templatePath='template.docx', outputPath='output.docx', replacements={'name':'John'})
- Convert format: word_file(operation='convert', path='doc.docx', outputPath='doc.pdf', format='pdf')
- Merge documents: word_file(operation='merge', inputPaths=['doc1.docx','doc2.docx'], outputPath='merged.docx')
- Split document: word_file(operation='split', path='doc.docx', outputDir='output/', splitBy='page')";

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
            replacements = new
            {
                type = "object",
                description = "Key-value pairs for placeholder replacements (for create_from_template)"
            },
            placeholderStyle = new
            {
                type = "string",
                description = "Placeholder format: doubleCurly, singleCurly, square (default: doubleCurly)",
                @enum = new[] { "doubleCurly", "singleCurly", "square" }
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
                    // Use predefined paper size
                    switch (paperSize.ToUpper())
                    {
                        case "A4":
                            pageSetup.PageWidth = 595.3; // 21.0 cm
                            pageSetup.PageHeight = 841.9; // 29.7 cm
                            break;
                        case "LETTER":
                            pageSetup.PageWidth = 612; // 8.5 inch
                            pageSetup.PageHeight = 792; // 11 inch
                            break;
                        case "A3":
                            pageSetup.PageWidth = 841.9; // 29.7 cm
                            pageSetup.PageHeight = 1190.55; // 42.0 cm
                            break;
                        case "LEGAL":
                            pageSetup.PageWidth = 612; // 8.5 inch
                            pageSetup.PageHeight = 1008; // 14 inch
                            break;
                        default:
                            pageSetup.PageWidth = 595.3; // Default to A4
                            pageSetup.PageHeight = 841.9;
                            break;
                    }
                }
                else
                {
                    // Use custom dimensions if specified
                    pageSetup.PageWidth = pageWidth ?? 595.3; // Default to A4 width

                    pageSetup.PageHeight = pageHeight ?? 841.9; // Default to A4 height
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
    ///     Creates a document from a template
    /// </summary>
    /// <param name="arguments">JSON arguments containing templatePath, outputPath, optional data</param>
    /// <returns>Success message with output path</returns>
    private Task<string> CreateFromTemplate(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var templatePath = ArgumentHelper.GetString(arguments, "templatePath");
            var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
            var placeholderStyle = ArgumentHelper.GetString(arguments, "placeholderStyle", "doubleCurly");

            SecurityHelper.ValidateFilePath(templatePath, "templatePath", true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            var replacements = new Dictionary<string, string>();
            if (arguments?.ContainsKey("replacements") == true)
            {
                var replacementsObj = arguments["replacements"]?.AsObject();
                if (replacementsObj != null)
                    foreach (var kvp in replacementsObj)
                    {
                        var key = kvp.Key;
                        var value = kvp.Value?.GetValue<string>() ?? "";

                        if (!IsValidPlaceholder(key, placeholderStyle))
                            key = FormatPlaceholder(key, placeholderStyle);

                        replacements[key] = value;
                    }
            }

            if (replacements.Count == 0)
                throw new ArgumentException("replacements cannot be empty");

            var doc = new Document(templatePath);

            foreach (var kvp in replacements)
                doc.Range.Replace(kvp.Key, kvp.Value, new FindReplaceOptions());

            doc.Save(outputPath);
            return $"Document created from template: {outputPath} (replaced {replacements.Count} placeholders)";
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
            var format = ArgumentHelper.GetString(arguments, "format").ToLower();

            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

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

            SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => p != null).ToList();

            if (inputPaths.Count == 0)
                throw new ArgumentException("At least one input path is required");

            foreach (var inputPath in inputPaths) SecurityHelper.ValidateFilePath(inputPath!, "inputPaths", true);

            var mergedDoc = new Document(inputPaths[0]);

            for (var i = 1; i < inputPaths.Count; i++)
            {
                var doc = new Document(inputPaths[i]);
                mergedDoc.AppendDocument(doc, ImportFormatMode.KeepSourceFormatting);
            }

            mergedDoc.Save(outputPath);
            return $"Merged {inputPaths.Count} documents into: {outputPath}";
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
                    var sectionDoc = new Document();
                    sectionDoc.FirstSection.Remove();
                    var importedSection = sectionDoc.ImportNode(doc.Sections[i], true);
                    sectionDoc.AppendChild(importedSection);

                    var outputPath = Path.Combine(outputDir, $"{fileBaseName}_section_{i + 1}.docx");
                    sectionDoc.Save(outputPath);
                }

                return $"Document split into {doc.Sections.Count} sections in: {outputDir}";
            }

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

    /// <summary>
    ///     Checks if a placeholder key matches the expected format for the given style
    /// </summary>
    /// <param name="key">Placeholder key to validate</param>
    /// <param name="style">Placeholder style (doubleCurly, singleCurly, square)</param>
    /// <returns>True if the key matches the expected format</returns>
    private bool IsValidPlaceholder(string key, string style)
    {
        return style.ToLower() switch
        {
            "singlecurly" => key.StartsWith("{") && key.EndsWith("}"),
            "square" => key.StartsWith("[") && key.EndsWith("]"),
            _ => key.StartsWith("{{") && key.EndsWith("}}")
        };
    }

    /// <summary>
    ///     Formats a placeholder key according to the specified style
    /// </summary>
    /// <param name="key">Placeholder key (may already contain brackets)</param>
    /// <param name="style">Placeholder style (doubleCurly, singleCurly, square)</param>
    /// <returns>Formatted placeholder string</returns>
    private string FormatPlaceholder(string key, string style)
    {
        key = key.Trim('{', '}', '[', ']');
        return style.ToLower() switch
        {
            "singlecurly" => $"{{{key}}}",
            "square" => $"[{key}]",
            _ => $"{{{{{key}}}}}"
        };
    }
}