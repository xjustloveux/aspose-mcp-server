using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Layout;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for page operations in Word documents
///     Merges: WordSetPageMarginsTool, WordSetPageOrientationTool, WordSetPageSizeTool,
///     WordSetPageNumberTool, WordSetPageSetupTool, WordDeletePageTool, WordInsertBlankPageTool, WordAddPageBreakTool
/// </summary>
public class WordPageTool : IAsposeTool
{
    public string Description =>
        @"Manage page settings in Word documents. Supports 8 operations: set_margins, set_orientation, set_size, set_page_number, set_page_setup, delete_page, insert_blank_page, add_page_break.

Usage examples:
- Set margins: word_page(operation='set_margins', path='doc.docx', top=72, bottom=72, left=72, right=72)
- Set orientation: word_page(operation='set_orientation', path='doc.docx', orientation='landscape')
- Set page size: word_page(operation='set_size', path='doc.docx', width=792, height=612)
- Set page number: word_page(operation='set_page_number', path='doc.docx', startNumber=1)
- Delete page: word_page(operation='delete_page', path='doc.docx', pageIndex=1)
- Insert blank page: word_page(operation='insert_blank_page', path='doc.docx', insertAtPageIndex=2)
- Add page break: word_page(operation='add_page_break', path='doc.docx', paragraphIndex=10)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set_margins': Set page margins (required params: path)
- 'set_orientation': Set page orientation (required params: path, orientation)
- 'set_size': Set page size (required params: path, width, height)
- 'set_page_number': Set page number format (required params: path, startNumber)
- 'set_page_setup': Set page setup (required params: path)
- 'delete_page': Delete a page (required params: path, pageIndex)
- 'insert_blank_page': Insert blank page (required params: path, insertAtParagraphIndex)
- 'add_page_break': Add page break (required params: path; optional: paragraphIndex)",
                @enum = new[]
                {
                    "set_margins", "set_orientation", "set_size", "set_page_number", "set_page_setup", "delete_page",
                    "insert_blank_page", "add_page_break"
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
            // Set margins parameters (72 points = 1 inch = 2.54 cm)
            top = new
            {
                type = "number",
                description = "Top margin in points (72 pts = 1 inch, e.g., 72 for 1 inch margin)"
            },
            bottom = new
            {
                type = "number",
                description = "Bottom margin in points (72 pts = 1 inch)"
            },
            left = new
            {
                type = "number",
                description = "Left margin in points (72 pts = 1 inch)"
            },
            right = new
            {
                type = "number",
                description = "Right margin in points (72 pts = 1 inch)"
            },
            // Set orientation parameters
            orientation = new
            {
                type = "string",
                description = "Orientation: Portrait or Landscape (required for set_orientation operation)",
                @enum = new[] { "Portrait", "Landscape" }
            },
            // Set size parameters (72 points = 1 inch = 2.54 cm)
            width = new
            {
                type = "number",
                description = "Page width in points (72 pts = 1 inch, e.g., A4 = 595 pts)"
            },
            height = new
            {
                type = "number",
                description = "Page height in points (72 pts = 1 inch, e.g., A4 = 842 pts)"
            },
            paperSize = new
            {
                type = "string",
                description =
                    "Predefined paper size: A4, Letter, Legal, A3, A5 (optional, overrides width/height, for set_size operation)",
                @enum = new[] { "A4", "Letter", "Legal", "A3", "A5" }
            },
            // Set page number parameters
            pageNumberFormat = new
            {
                type = "string",
                description = "Page number format: arabic, roman, letter (optional, for set_page_number operation)",
                @enum = new[] { "arabic", "roman", "letter" }
            },
            startingPageNumber = new
            {
                type = "number",
                description = "Starting page number (optional, for set_page_number operation)"
            },
            // Common parameters
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided applies to all sections)"
            },
            sectionIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of section indices (0-based, optional, overrides sectionIndex)"
            },
            // Delete page parameters
            pageIndex = new
            {
                type = "number",
                description = "Page index to delete (0-based, required for delete_page operation)"
            },
            // Insert blank page parameters
            insertAtPageIndex = new
            {
                type = "number",
                description = "Page index to insert blank page at (0-based, optional, for insert_blank_page operation)"
            },
            // Add page break parameters
            paragraphIndex = new
            {
                type = "number",
                description =
                    "Paragraph index to insert page break after (0-based, optional for add_page_break, defaults to document end)"
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
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation switch
        {
            "set_margins" => await SetMarginsAsync(path, outputPath, arguments),
            "set_orientation" => await SetOrientationAsync(path, outputPath, arguments),
            "set_size" => await SetSizeAsync(path, outputPath, arguments),
            "set_page_number" => await SetPageNumberAsync(path, outputPath, arguments),
            "set_page_setup" => await SetPageSetupAsync(path, outputPath, arguments),
            "delete_page" => await DeletePageAsync(path, outputPath, arguments),
            "insert_blank_page" => await InsertBlankPageAsync(path, outputPath, arguments),
            "add_page_break" => await AddPageBreakAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets target section indices based on sectionIndex or sectionIndices parameters
    /// </summary>
    /// <param name="doc">Word document</param>
    /// <param name="arguments">JSON arguments containing optional sectionIndex or sectionIndices</param>
    /// <param name="validateRange">Whether to validate section indices are within range</param>
    /// <returns>List of section indices to process</returns>
    private static List<int> GetTargetSections(Document doc, JsonObject? arguments, bool validateRange = true)
    {
        var sectionIndicesArray = ArgumentHelper.GetArray(arguments, "sectionIndices", false);
        var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

        if (sectionIndicesArray is { Count: > 0 })
        {
            var indices = sectionIndicesArray
                .Select(s => s?.GetValue<int>())
                .Where(s => s.HasValue)
                .Select(s => s!.Value)
                .ToList();

            if (validateRange)
                foreach (var idx in indices)
                    if (idx < 0 || idx >= doc.Sections.Count)
                        throw new ArgumentException(
                            $"sectionIndex {idx} must be between 0 and {doc.Sections.Count - 1}");

            return indices;
        }

        if (sectionIndex.HasValue)
        {
            if (validateRange && (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            return [sectionIndex.Value];
        }

        return Enumerable.Range(0, doc.Sections.Count).ToList();
    }

    /// <summary>
    ///     Sets page margins
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional top, bottom, left, right, sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> SetMarginsAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var top = ArgumentHelper.GetDoubleNullable(arguments, "top");
            var bottom = ArgumentHelper.GetDoubleNullable(arguments, "bottom");
            var left = ArgumentHelper.GetDoubleNullable(arguments, "left");
            var right = ArgumentHelper.GetDoubleNullable(arguments, "right");

            var doc = new Document(path);
            var sectionsToUpdate = GetTargetSections(doc, arguments);

            foreach (var idx in sectionsToUpdate)
            {
                var pageSetup = doc.Sections[idx].PageSetup;
                if (top.HasValue) pageSetup.TopMargin = top.Value;
                if (bottom.HasValue) pageSetup.BottomMargin = bottom.Value;
                if (left.HasValue) pageSetup.LeftMargin = left.Value;
                if (right.HasValue) pageSetup.RightMargin = right.Value;
            }

            doc.Save(outputPath);
            return $"Page margins updated for {sectionsToUpdate.Count} section(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Sets page orientation
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing orientation, optional sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> SetOrientationAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var orientation = ArgumentHelper.GetString(arguments, "orientation");

            var doc = new Document(path);
            var orientationEnum = orientation.ToLower() == "landscape" ? Orientation.Landscape : Orientation.Portrait;
            var sectionsToUpdate = GetTargetSections(doc, arguments);

            foreach (var idx in sectionsToUpdate)
                doc.Sections[idx].PageSetup.Orientation = orientationEnum;

            doc.Save(outputPath);
            return $"Page orientation set to {orientation} for {sectionsToUpdate.Count} section(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Sets page size
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional width, height, paperSize, sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> SetSizeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
            var height = ArgumentHelper.GetDoubleNullable(arguments, "height");
            var paperSize = ArgumentHelper.GetStringNullable(arguments, "paperSize");

            var doc = new Document(path);
            var sectionsToUpdate = GetTargetSections(doc, arguments);

            foreach (var idx in sectionsToUpdate)
            {
                var pageSetup = doc.Sections[idx].PageSetup;

                if (!string.IsNullOrEmpty(paperSize))
                {
                    pageSetup.PaperSize = paperSize.ToUpper() switch
                    {
                        "A4" => PaperSize.A4,
                        "LETTER" => PaperSize.Letter,
                        "LEGAL" => PaperSize.Legal,
                        "A3" => PaperSize.A3,
                        "A5" => PaperSize.A5,
                        _ => PaperSize.A4
                    };
                }
                else if (width.HasValue && height.HasValue)
                {
                    pageSetup.PageWidth = width.Value;
                    pageSetup.PageHeight = height.Value;
                }
                else
                {
                    throw new ArgumentException("Either paperSize or both width and height must be provided");
                }
            }

            doc.Save(outputPath);
            return $"Page size updated for {sectionsToUpdate.Count} section(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Sets page numbering
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional startNumber, numberFormat, sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> SetPageNumberAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageNumberFormat = ArgumentHelper.GetStringNullable(arguments, "pageNumberFormat");
            var startingPageNumber = ArgumentHelper.GetIntNullable(arguments, "startingPageNumber");
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);
            List<int> sectionsToUpdate;

            if (sectionIndex.HasValue)
            {
                if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                    throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
                sectionsToUpdate = [sectionIndex.Value];
            }
            else
            {
                sectionsToUpdate = [0];
            }

            foreach (var idx in sectionsToUpdate)
            {
                var pageSetup = doc.Sections[idx].PageSetup;

                if (!string.IsNullOrEmpty(pageNumberFormat))
                {
                    var numStyle = pageNumberFormat.ToLower() switch
                    {
                        "roman" => NumberStyle.UppercaseRoman,
                        "letter" => NumberStyle.UppercaseLetter,
                        _ => NumberStyle.Arabic
                    };
                    pageSetup.PageNumberStyle = numStyle;
                }

                if (startingPageNumber.HasValue)
                {
                    pageSetup.RestartPageNumbering = true;
                    pageSetup.PageStartingNumber = startingPageNumber.Value;
                }
            }

            doc.Save(outputPath);
            return $"Page number settings updated for {sectionsToUpdate.Count} section(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Sets page setup properties
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing various page setup options, optional sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> SetPageSetupAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            // This is a combined operation that can set multiple page setup properties
            var doc = new Document(path);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var pageSetup = doc.Sections[sectionIndex].PageSetup;
            var changes = new List<string>();

            // Apply all page setup parameters
            var top = ArgumentHelper.GetDoubleNullable(arguments, "top");
            if (top.HasValue)
            {
                pageSetup.TopMargin = top.Value;
                changes.Add($"Top margin: {top.Value}");
            }

            var bottom = ArgumentHelper.GetDoubleNullable(arguments, "bottom");
            if (bottom.HasValue)
            {
                pageSetup.BottomMargin = bottom.Value;
                changes.Add($"Bottom margin: {bottom.Value}");
            }

            var left = ArgumentHelper.GetDoubleNullable(arguments, "left");
            if (left.HasValue)
            {
                pageSetup.LeftMargin = left.Value;
                changes.Add($"Left margin: {left.Value}");
            }

            var right = ArgumentHelper.GetDoubleNullable(arguments, "right");
            if (right.HasValue)
            {
                pageSetup.RightMargin = right.Value;
                changes.Add($"Right margin: {right.Value}");
            }

            var orientation = ArgumentHelper.GetStringNullable(arguments, "orientation");
            if (!string.IsNullOrEmpty(orientation))
            {
                pageSetup.Orientation =
                    orientation.ToLower() == "landscape" ? Orientation.Landscape : Orientation.Portrait;
                changes.Add($"Orientation: {orientation}");
            }

            doc.Save(outputPath);
            return $"Page setup updated: {string.Join(", ", changes)}";
        });
    }

    /// <summary>
    ///     Deletes a page from the document using ExtractPages
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeletePageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

            var doc = new Document(path);
            var pageCount = doc.PageCount;

            if (pageIndex < 0 || pageIndex >= pageCount)
                throw new ArgumentException(
                    $"pageIndex must be between 0 and {pageCount - 1} (document has {pageCount} pages)");

            // Use ExtractPages to rebuild document without the specified page
            var resultDoc = new Document();
            resultDoc.RemoveAllChildren();

            // Extract pages before the deleted page
            if (pageIndex > 0)
            {
                var beforePages = doc.ExtractPages(0, pageIndex);
                foreach (var section in beforePages.Sections.Cast<Section>())
                    resultDoc.AppendChild(resultDoc.ImportNode(section, true));
            }

            // Extract pages after the deleted page
            if (pageIndex < pageCount - 1)
            {
                var afterPages = doc.ExtractPages(pageIndex + 1, pageCount - pageIndex - 1);
                foreach (var section in afterPages.Sections.Cast<Section>())
                    resultDoc.AppendChild(resultDoc.ImportNode(section, true));
            }

            resultDoc.Save(outputPath);
            return
                $"Page {pageIndex} deleted successfully (document now has {resultDoc.PageCount} pages)\nOutput: {outputPath}";
        });
    }

    /// <summary>
    ///     Inserts a blank page into the document using LayoutCollector for precise positioning
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional insertAtPageIndex</param>
    /// <returns>Success message</returns>
    private Task<string> InsertBlankPageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var insertAtPageIndex = ArgumentHelper.GetIntNullable(arguments, "insertAtPageIndex");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            if (insertAtPageIndex is > 0)
            {
                var pageCount = doc.PageCount;
                if (insertAtPageIndex.Value > pageCount)
                    throw new ArgumentException(
                        $"insertAtPageIndex must be between 0 and {pageCount} (document has {pageCount} pages)");

                // Use LayoutCollector to find the first node on the target page
                var layoutCollector = new LayoutCollector(doc);
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

                Paragraph? targetParagraph = null;
                foreach (var para in paragraphs)
                {
                    var paraPage = layoutCollector.GetStartPageIndex(para);
                    if (paraPage == insertAtPageIndex.Value + 1) // LayoutCollector uses 1-based page index
                    {
                        targetParagraph = para;
                        break;
                    }
                }

                if (targetParagraph != null)
                {
                    builder.MoveTo(targetParagraph);
                    builder.InsertBreak(BreakType.PageBreak);
                }
                else
                {
                    // Fallback: insert at document end
                    builder.MoveToDocumentEnd();
                    builder.InsertBreak(BreakType.PageBreak);
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.PageBreak);
            }

            doc.Save(outputPath);
            return $"Blank page inserted at page {insertAtPageIndex ?? doc.PageCount}\nOutput: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds a page break to the document at specified paragraph or document end
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional paragraphIndex</param>
    /// <returns>Success message</returns>
    private Task<string> AddPageBreakAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                    throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");

                builder.MoveTo(paragraphs[paragraphIndex.Value]);
                builder.InsertBreak(BreakType.PageBreak);
            }
            else
            {
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.PageBreak);
            }

            doc.Save(outputPath);
            var location = paragraphIndex.HasValue ? $"after paragraph {paragraphIndex.Value}" : "at document end";
            return $"Page break added {location}\nOutput: {outputPath}";
        });
    }
}