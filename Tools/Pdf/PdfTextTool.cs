using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing text in PDF documents (add, edit, extract)
/// </summary>
public class PdfTextTool : IAsposeTool
{
    public string Description => @"Manage text in PDF documents. Supports 3 operations: add, edit, extract.

Usage examples:
- Add text: pdf_text(operation='add', path='doc.pdf', pageIndex=1, text='Hello World', x=100, y=700)
- Add text with font: pdf_text(operation='add', path='doc.pdf', pageIndex=1, text='Hello', x=100, y=700, fontName='Arial', fontSize=14)
- Edit text: pdf_text(operation='edit', path='doc.pdf', pageIndex=1, oldText='old', newText='new')
- Edit all occurrences: pdf_text(operation='edit', path='doc.pdf', pageIndex=1, oldText='old', newText='new', replaceAll=true)
- Extract text: pdf_text(operation='extract', path='doc.pdf', pageIndex=1)
- Extract with font info: pdf_text(operation='extract', path='doc.pdf', pageIndex=1, includeFontInfo=true)
- Extract raw text: pdf_text(operation='extract', path='doc.pdf', pageIndex=1, extractionMode='raw')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add text to page (required params: path, pageIndex, text)
- 'edit': Edit text on page (required params: path, pageIndex, text)
- 'extract': Extract text from page (required params: path, pageIndex, outputPath)",
                @enum = new[] { "add", "edit", "extract" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, defaults to overwrite input for add/edit, required for extract)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add, edit, extract)"
            },
            text = new
            {
                type = "string",
                description = "Text to add (required for add)"
            },
            x = new
            {
                type = "number",
                description = "X position in PDF coordinates, origin at bottom-left corner (for add, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position in PDF coordinates, origin at bottom-left corner (for add, default: 700)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (for add, default: 'Arial')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (for add, default: 12)"
            },
            oldText = new
            {
                type = "string",
                description = "Text to replace (required for edit)"
            },
            newText = new
            {
                type = "string",
                description = "New text (required for edit)"
            },
            replaceAll = new
            {
                type = "boolean",
                description = "Replace all occurrences (for edit, default: false)"
            },
            includeFontInfo = new
            {
                type = "boolean",
                description = "Include font information (for extract, default: false)"
            },
            extractionMode = new
            {
                type = "string",
                description =
                    "Text extraction mode (for extract, default: 'pure'). 'pure' preserves formatting, 'raw' extracts raw text without formatting",
                @enum = new[] { "pure", "raw" }
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        // Only get outputPath for operations that modify the document
        string? outputPath = null;
        if (operation.ToLower() != "extract")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddText(path, outputPath!, arguments),
            "edit" => await EditText(path, outputPath!, arguments),
            "extract" => await ExtractText(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds text to a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, text, optional x, y, fontName, fontSize</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range</exception>
    private Task<string> AddText(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var text = ArgumentHelper.GetString(arguments, "text");
            var x = ArgumentHelper.GetDouble(arguments, "x", "x", false, 100);
            var y = ArgumentHelper.GetDouble(arguments, "y", "y", false, 700);
            var fontName = ArgumentHelper.GetString(arguments, "fontName", "Arial");
            var fontSize = ArgumentHelper.GetDouble(arguments, "fontSize", "fontSize", false, 12);

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            var textFragment = new TextFragment(text)
            {
                Position = new Position(x, y)
            };

            // Apply font settings using FontHelper
            FontHelper.Pdf.ApplyFontSettings(
                textFragment.TextState,
                fontName,
                fontSize
            );

            var textBuilder = new TextBuilder(page);
            textBuilder.AppendText(textFragment);
            document.Save(outputPath);
            return $"Added text to page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits text on a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, oldText, newText, optional replaceAll</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range or text not found</exception>
    private Task<string> EditText(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var oldText = ArgumentHelper.GetString(arguments, "oldText");
            var newText = ArgumentHelper.GetString(arguments, "newText");
            var replaceAll = ArgumentHelper.GetBool(arguments, "replaceAll", false);

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];

            try
            {
                // Try exact match first
                var textFragmentAbsorber = new TextFragmentAbsorber(oldText);
                page.Accept(textFragmentAbsorber);

                var fragments = textFragmentAbsorber.TextFragments;
                var normalizedOldText = Regex.Replace(oldText, @"\s+", " ").Trim();

                // If exact match fails, try with normalized whitespace
                if (fragments.Count == 0 && normalizedOldText != oldText)
                {
                    textFragmentAbsorber = new TextFragmentAbsorber(normalizedOldText);
                    page.Accept(textFragmentAbsorber);
                    fragments = textFragmentAbsorber.TextFragments;
                }

                // If still no match, try case-insensitive search and partial matching
                if (fragments.Count == 0)
                {
                    var allTextAbsorber = new TextFragmentAbsorber();
                    page.Accept(allTextAbsorber);
                    var allFragments = allTextAbsorber.TextFragments;
                    var matchingFragments = new List<TextFragment>();

                    // Clean oldText: remove null characters and normalize
                    var cleanedOldText = oldText.Replace("\u0000", "").Trim();
                    var normalizedCleanedOldText = Regex.Replace(cleanedOldText, @"\s+", " ").Trim();

                    foreach (var fragment in allFragments)
                        try
                        {
                            var fragmentText = (fragment.Text?.Replace("\u0000", "") ?? "").Trim();

                            if (IsTextMatch(fragmentText, oldText, normalizedOldText, cleanedOldText,
                                    normalizedCleanedOldText))
                                matchingFragments.Add(fragment);
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine(
                                $"[WARN] Error processing text fragment during search: {ex.Message}");
                        }

                    if (matchingFragments.Count > 0)
                    {
                        var replaceCount = replaceAll ? matchingFragments.Count : 1;
                        for (var i = 0; i < replaceCount && i < matchingFragments.Count; i++)
                            try
                            {
                                matchingFragments[i].Text = newText;
                            }
                            catch (Exception ex)
                            {
                                throw new ArgumentException(
                                    $"Failed to replace text at occurrence {i + 1}: {ex.Message}");
                            }

                        document.Save(outputPath);
                        return
                            $"Replaced {replaceCount} occurrence(s) of '{oldText}' on page {pageIndex}. Output: {outputPath}";
                    }
                }

                if (fragments.Count == 0)
                {
                    // Provide helpful error message with actual extracted text
                    var textAbsorber = new TextAbsorber();
                    page.Accept(textAbsorber);
                    var pageText = textAbsorber.Text ?? "";

                    // Also try to find similar text using fuzzy matching
                    var normalizedPageText = Regex.Replace(pageText, @"\s+", " ").Trim();
                    var normalizedSearchText = Regex.Replace(oldText, @"\s+", " ").Trim();
                    var cleanedPageText = pageText.Replace("\u0000", "").Trim();
                    var cleanedSearchText = oldText.Replace("\u0000", "").Trim();

                    // Get fragment details for debugging
                    var fragmentAbsorber = new TextFragmentAbsorber();
                    page.Accept(fragmentAbsorber);
                    var fragmentDetails = new List<string>();
                    foreach (var frag in fragmentAbsorber.TextFragments)
                    {
                        var fragText = (frag.Text?.Replace("\u0000", "") ?? "").Trim();
                        if (!string.IsNullOrEmpty(fragText)) fragmentDetails.Add($"'{fragText}'");
                    }

                    var preview = pageText.Length > 200 ? pageText.Substring(0, 200) + "..." : pageText;
                    var errorMsg = $"Text '{oldText}' not found on page {pageIndex}.";

                    // Check if normalized versions match (ignoring whitespace differences)
                    if (normalizedPageText.Contains(normalizedSearchText, StringComparison.OrdinalIgnoreCase))
                        errorMsg +=
                            " Note: The text exists but with different whitespace. Try matching the exact extracted text format.";
                    else if (cleanedPageText.Contains(cleanedSearchText, StringComparison.OrdinalIgnoreCase))
                        errorMsg +=
                            " Note: The text exists but contains null characters (\\u0000). Try using the cleaned extracted text.";

                    if (fragmentDetails.Count > 0)
                    {
                        errorMsg +=
                            $" Found {fragmentDetails.Count} text fragment(s): {string.Join(", ", fragmentDetails.Take(5))}";
                        if (fragmentDetails.Count > 5) errorMsg += "...";
                    }

                    errorMsg += $" Page text preview: {preview}";
                    throw new ArgumentException(errorMsg);
                }

                // TextFragmentCollection is 1-based, so we need to use 1-based indexing
                // However, we should iterate through the collection properly
                var finalReplaceCount = replaceAll ? fragments.Count : 1;
                var replacedCount = 0;

                // Use foreach to iterate through fragments (safer than index-based access)
                foreach (var fragment in fragments)
                {
                    if (replacedCount >= finalReplaceCount)
                        break;

                    try
                    {
                        fragment.Text = newText;
                        replacedCount++;
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException(
                            $"Failed to replace text at occurrence {replacedCount + 1}: {ex.Message}");
                    }
                }

                if (replacedCount == 0 && fragments.Count > 0)
                    throw new ArgumentException(
                        $"Failed to replace any text fragments. Found {fragments.Count} fragment(s) but replacement failed.");

                document.Save(outputPath);
                return
                    $"Replaced {finalReplaceCount} occurrence(s) of '{oldText}' on page {pageIndex}. Output: {outputPath}";
            }
            catch (ArgumentException)
            {
                // Re-throw ArgumentException as-is
                throw;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Failed to edit text on page {pageIndex}: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Extracts text from a PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, optional includeFontInfo, extractionMode</param>
    /// <returns>JSON formatted string with extracted text</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range</exception>
    private Task<string> ExtractText(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var includeFontInfo = ArgumentHelper.GetBool(arguments, "includeFontInfo", false);
            var extractionMode = ArgumentHelper.GetString(arguments, "extractionMode", "pure").ToLower();

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];

            // Configure text extraction options based on mode
            var textAbsorber = new TextAbsorber();
            if (extractionMode == "raw")
                textAbsorber.ExtractionOptions =
                    new TextExtractionOptions(TextExtractionOptions.TextFormattingMode.Raw);

            page.Accept(textAbsorber);

            if (includeFontInfo)
            {
                var textFragmentAbsorber = new TextFragmentAbsorber();
                page.Accept(textFragmentAbsorber);
                var fragments = new List<object>();

                foreach (var fragment in textFragmentAbsorber.TextFragments)
                    fragments.Add(new
                    {
                        text = fragment.Text,
                        fontName = fragment.TextState.Font.FontName,
                        fontSize = fragment.TextState.FontSize
                    });

                var result = new
                {
                    pageIndex,
                    totalPages = document.Pages.Count,
                    fragmentCount = fragments.Count,
                    fragments
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            else
            {
                var result = new
                {
                    pageIndex,
                    totalPages = document.Pages.Count,
                    text = textAbsorber.Text
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
        });
    }

    /// <summary>
    ///     Checks if a text fragment matches the search text using multiple strategies
    /// </summary>
    /// <param name="fragmentText">The text from the PDF fragment</param>
    /// <param name="searchText">The text to search for</param>
    /// <param name="normalizedSearchText">Normalized version of the search text (whitespace collapsed)</param>
    /// <param name="cleanedSearchText">Cleaned version of the search text (null chars removed)</param>
    /// <param name="normalizedCleanedSearchText">Normalized and cleaned version of the search text</param>
    /// <returns>True if the fragment matches the search text</returns>
    private static bool IsTextMatch(
        string fragmentText,
        string searchText,
        string normalizedSearchText,
        string cleanedSearchText,
        string normalizedCleanedSearchText)
    {
        if (string.IsNullOrEmpty(fragmentText)) return false;

        var normalizedFragmentText = Regex.Replace(fragmentText, @"\s+", " ").Trim();

        // Strategy 1: Exact match (case-insensitive)
        if (fragmentText.Equals(searchText, StringComparison.OrdinalIgnoreCase) ||
            fragmentText.Equals(cleanedSearchText, StringComparison.OrdinalIgnoreCase))
            return true;

        // Strategy 2: Normalized match (case-insensitive)
        if (normalizedFragmentText.Equals(searchText, StringComparison.OrdinalIgnoreCase) ||
            normalizedFragmentText.Equals(normalizedSearchText, StringComparison.OrdinalIgnoreCase) ||
            normalizedFragmentText.Equals(normalizedCleanedSearchText, StringComparison.OrdinalIgnoreCase))
            return true;

        // Strategy 3: Partial match - search text contains fragment
        if (searchText.Contains(fragmentText, StringComparison.OrdinalIgnoreCase) ||
            cleanedSearchText.Contains(fragmentText, StringComparison.OrdinalIgnoreCase))
            return true;

        // Strategy 4: Partial match - fragment contains search text
        if (fragmentText.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
            fragmentText.Contains(cleanedSearchText, StringComparison.OrdinalIgnoreCase))
            return true;

        // Strategy 5: Normalized partial match
        if (normalizedFragmentText.Length > 0 &&
            (normalizedFragmentText.Contains(normalizedSearchText, StringComparison.OrdinalIgnoreCase) ||
             normalizedFragmentText.Contains(normalizedCleanedSearchText, StringComparison.OrdinalIgnoreCase)))
            return true;

        return false;
    }
}