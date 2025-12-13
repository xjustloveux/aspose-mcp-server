using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordFileTool : IAsposeTool
{
    public string Description => @"Perform file operations on Word documents. Supports 5 operations: create, create_from_template, convert, merge, split.

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
                description = "Output file path (required for create, create_from_template, convert, and merge operations)"
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
                description = "Predefined paper size: A4, Letter, A3, Legal (for create, default: A4). Overrides pageWidth/pageHeight if specified.",
                @enum = new[] { "A4", "Letter", "A3", "Legal" }
            },
            pageWidth = new
            {
                type = "number",
                description = "Page width in points (for create, e.g., 595.3 for A4, 612 for Letter). Default: 595.3 (A4)"
            },
            pageHeight = new
            {
                type = "number",
                description = "Page height in points (for create, e.g., 841.9 for A4, 792 for Letter). Default: 841.9 (A4)"
            },
            headerDistance = new
            {
                type = "number",
                description = "Header distance from page top in points (for create, e.g., 45.35 pt = 1.6 cm). Default: 35.4"
            },
            footerDistance = new
            {
                type = "number",
                description = "Footer distance from page bottom in points (for create, e.g., 45.35 pt = 1.6 cm). Default: 35.4"
            }
        },
        required = new[] { "operation" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

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

    private async Task<string> CreateDocument(JsonObject? arguments)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var content = arguments?["content"]?.GetValue<string>();
        var skipInitialContent = arguments?["skipInitialContent"]?.GetValue<bool>() ?? false;
        var marginTop = arguments?["marginTop"]?.GetValue<double?>() ?? 70.87;
        var marginBottom = arguments?["marginBottom"]?.GetValue<double?>() ?? 70.87;
        var marginLeft = arguments?["marginLeft"]?.GetValue<double?>() ?? 70.87;
        var marginRight = arguments?["marginRight"]?.GetValue<double?>() ?? 70.87;
        var compatibilityMode = arguments?["compatibilityMode"]?.GetValue<string>() ?? "Word2019";
        var pageWidth = arguments?["pageWidth"]?.GetValue<double?>();
        var pageHeight = arguments?["pageHeight"]?.GetValue<double?>();
        var paperSize = arguments?["paperSize"]?.GetValue<string>() ?? "A4";
        var headerDistance = arguments?["headerDistance"]?.GetValue<double?>() ?? 35.4;
        var footerDistance = arguments?["footerDistance"]?.GetValue<double?>() ?? 35.4;

        var doc = new Document();
        
        var wordVersion = compatibilityMode switch
        {
            "Word2019" => Aspose.Words.Settings.MsWordVersion.Word2019,
            "Word2016" => Aspose.Words.Settings.MsWordVersion.Word2016,
            "Word2013" => Aspose.Words.Settings.MsWordVersion.Word2013,
            "Word2010" => Aspose.Words.Settings.MsWordVersion.Word2010,
            "Word2007" => Aspose.Words.Settings.MsWordVersion.Word2007,
            _ => Aspose.Words.Settings.MsWordVersion.Word2019
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
                        pageSetup.PageWidth = 595.3;  // 21.0 cm
                        pageSetup.PageHeight = 841.9; // 29.7 cm
                        break;
                    case "LETTER":
                        pageSetup.PageWidth = 612;    // 8.5 inch
                        pageSetup.PageHeight = 792;   // 11 inch
                        break;
                    case "A3":
                        pageSetup.PageWidth = 841.9;  // 29.7 cm
                        pageSetup.PageHeight = 1190.55; // 42.0 cm
                        break;
                    case "LEGAL":
                        pageSetup.PageWidth = 612;    // 8.5 inch
                        pageSetup.PageHeight = 1008;  // 14 inch
                        break;
                    default:
                        pageSetup.PageWidth = 595.3;  // Default to A4
                        pageSetup.PageHeight = 841.9;
                        break;
                }
            }
            else
            {
                // Use custom dimensions if specified
                if (pageWidth.HasValue)
                    pageSetup.PageWidth = pageWidth.Value;
                else
                    pageSetup.PageWidth = 595.3; // Default to A4 width

                if (pageHeight.HasValue)
                    pageSetup.PageHeight = pageHeight.Value;
                else
                    pageSetup.PageHeight = 841.9; // Default to A4 height
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
            if (doc.FirstSection != null && doc.FirstSection.Body != null)
            {
                doc.FirstSection.Body.RemoveAllChildren();
                var firstPara = new Paragraph(doc);
                firstPara.ParagraphFormat.SpaceBefore = 0;
                firstPara.ParagraphFormat.SpaceAfter = 0;
                firstPara.ParagraphFormat.LineSpacing = 12;
                doc.FirstSection.Body.AppendChild(firstPara);
            }
        }
        else if (!string.IsNullOrEmpty(content))
        {
            builder.Write(content);
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Word document created successfully at: {outputPath}");
    }

    private async Task<string> CreateFromTemplate(JsonObject? arguments)
    {
        var templatePath = arguments?["templatePath"]?.GetValue<string>() ?? throw new ArgumentException("templatePath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var placeholderStyle = arguments?["placeholderStyle"]?.GetValue<string>() ?? "doubleCurly";

        SecurityHelper.ValidateFilePath(templatePath, "templatePath");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"Template file not found: {templatePath}");

        var replacements = new Dictionary<string, string>();
        if (arguments?.ContainsKey("replacements") == true)
        {
            var replacementsObj = arguments["replacements"]?.AsObject();
            if (replacementsObj != null)
            {
                foreach (var kvp in replacementsObj)
                {
                    var key = kvp.Key;
                    var value = kvp.Value?.GetValue<string>() ?? "";
                    
                    if (!IsValidPlaceholder(key, placeholderStyle))
                        key = FormatPlaceholder(key, placeholderStyle);
                    
                    replacements[key] = value;
                }
            }
        }

        if (replacements.Count == 0)
            throw new ArgumentException("replacements cannot be empty");

        var doc = new Document(templatePath);

        foreach (var kvp in replacements)
            doc.Range.Replace(kvp.Key, kvp.Value, new Aspose.Words.Replacing.FindReplaceOptions());

        doc.Save(outputPath);
        return await Task.FromResult($"Document created from template: {outputPath} (replaced {replacements.Count} placeholders)");
    }

    private async Task<string> ConvertDocument(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var format = arguments?["format"]?.GetValue<string>()?.ToLower() ?? throw new ArgumentException("format is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

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
        return await Task.FromResult($"Document converted from {path} to {outputPath} ({format})");
    }

    private async Task<string> MergeDocuments(JsonObject? arguments)
    {
        var inputPathsArray = arguments?["inputPaths"]?.AsArray() ?? throw new ArgumentException("inputPaths is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => p != null).ToList();
        
        if (inputPaths.Count == 0)
            throw new ArgumentException("At least one input path is required");

        foreach (var inputPath in inputPaths)
        {
            SecurityHelper.ValidateFilePath(inputPath!, "inputPaths");
        }

        var mergedDoc = new Document(inputPaths[0]);

        for (int i = 1; i < inputPaths.Count; i++)
        {
            var doc = new Document(inputPaths[i]);
            mergedDoc.AppendDocument(doc, ImportFormatMode.KeepSourceFormatting);
        }

        mergedDoc.Save(outputPath);
        return await Task.FromResult($"Merged {inputPaths.Count} documents into: {outputPath}");
    }

    private async Task<string> SplitDocument(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? throw new ArgumentException("outputDir is required");
        var splitBy = arguments?["splitBy"]?.GetValue<string>() ?? "section";

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputDir, "outputDir");

        Directory.CreateDirectory(outputDir);

        var doc = new Document(path);
        var fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(path));

        if (splitBy.ToLower() == "section")
        {
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                var sectionDoc = new Document();
                sectionDoc.FirstSection.Remove();
                var importedSection = sectionDoc.ImportNode(doc.Sections[i], true);
                sectionDoc.AppendChild(importedSection);

                var outputPath = Path.Combine(outputDir, $"{fileBaseName}_section_{i + 1}.docx");
                sectionDoc.Save(outputPath);
            }

            return await Task.FromResult($"Document split into {doc.Sections.Count} sections in: {outputDir}");
        }
        else
        {
            var pageCount = doc.PageCount;
            for (int i = 0; i < pageCount; i++)
            {
                var pageDoc = doc.ExtractPages(i, 1);
                var outputPath = Path.Combine(outputDir, $"{fileBaseName}_page_{i + 1}.docx");
                pageDoc.Save(outputPath);
            }

            return await Task.FromResult($"Document split into {pageCount} pages in: {outputDir}");
        }
    }

    private bool IsValidPlaceholder(string key, string style)
    {
        return style.ToLower() switch
        {
            "singlecurly" => key.StartsWith("{") && key.EndsWith("}"),
            "square" => key.StartsWith("[") && key.EndsWith("]"),
            _ => key.StartsWith("{{") && key.EndsWith("}}")
        };
    }

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

