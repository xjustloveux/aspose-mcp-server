using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word sections (insert, delete, get info)
///     Merges: WordInsertSectionTool, WordDeleteSectionTool, WordGetSectionsTool, WordGetSectionsInfoTool
/// </summary>
public class WordSectionTool : IAsposeTool
{
    public string Description => @"Manage Word document sections. Supports 3 operations: insert, delete, get.

Usage examples:
- Insert section: word_section(operation='insert', path='doc.docx', sectionBreakType='NextPage', insertAtParagraphIndex=5)
- Delete section: word_section(operation='delete', path='doc.docx', sectionIndex=1)
- Get sections: word_section(operation='get', path='doc.docx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'insert': Insert a section break (required params: path, sectionBreakType)
- 'delete': Delete a section (required params: path, sectionIndex)
- 'get': Get all sections info (required params: path)",
                @enum = new[] { "insert", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for insert/delete operations)"
            },
            sectionBreakType = new
            {
                type = "string",
                description =
                    "Section break type: 'NextPage', 'Continuous', 'EvenPage', 'OddPage' (required for insert operation)",
                @enum = new[] { "NextPage", "Continuous", "EvenPage", "OddPage" }
            },
            insertAtParagraphIndex = new
            {
                type = "number",
                description =
                    "Paragraph index to insert section break after (0-based, optional, default: end of document, for insert operation)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, required for delete operation, optional for get operation)"
            },
            sectionIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description =
                    "Array of section indices to delete (0-based, optional, overrides sectionIndex, for delete operation)"
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

        return operation.ToLower() switch
        {
            "insert" => await InsertSectionAsync(arguments, path),
            "delete" => await DeleteSectionAsync(arguments, path),
            "get" => await GetSectionsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a section break into the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing sectionBreakType, optional insertAtParagraphIndex, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private Task<string> InsertSectionAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            var sectionBreakType = ArgumentHelper.GetString(arguments, "sectionBreakType");
            var insertAtParagraphIndex = ArgumentHelper.GetIntNullable(arguments, "insertAtParagraphIndex");
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            var breakType = sectionBreakType switch
            {
                "NextPage" => SectionStart.NewPage,
                "Continuous" => SectionStart.Continuous,
                "EvenPage" => SectionStart.EvenPage,
                "OddPage" => SectionStart.OddPage,
                _ => SectionStart.NewPage
            };

            if (insertAtParagraphIndex.HasValue)
            {
                if (insertAtParagraphIndex.Value == -1)
                {
                    // insertAtParagraphIndex=-1 means document end
                    builder.MoveToDocumentEnd();
                    builder.InsertBreak(BreakType.SectionBreakContinuous);
                    builder.CurrentSection.PageSetup.SectionStart = breakType;
                }
                else
                {
                    var actualSectionIndex = sectionIndex ?? 0;
                    if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count) actualSectionIndex = 0;

                    var section = doc.Sections[actualSectionIndex];
                    var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

                    if (insertAtParagraphIndex.Value < 0 || insertAtParagraphIndex.Value >= paragraphs.Count)
                        throw new ArgumentException(
                            $"insertAtParagraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

                    var para = paragraphs[insertAtParagraphIndex.Value];
                    builder.MoveTo(para);
                    builder.InsertBreak(BreakType.SectionBreakContinuous);
                    builder.CurrentSection.PageSetup.SectionStart = breakType;
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.SectionBreakContinuous);
                builder.CurrentSection.PageSetup.SectionStart = breakType;
            }

            doc.Save(outputPath);
            return $"Section break inserted ({sectionBreakType}): {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a section from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing sectionIndex, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteSectionAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
            var sectionIndicesArray = ArgumentHelper.GetArray(arguments, "sectionIndices", false);

            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);
            if (doc.Sections.Count <= 1)
                throw new ArgumentException("Cannot delete the last section. Document must have at least one section.");

            List<int> sectionsToDelete;
            if (sectionIndicesArray is { Count: > 0 })
                sectionsToDelete = sectionIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue)
                    .Select(s => s!.Value).OrderByDescending(s => s).ToList();
            else if (sectionIndex.HasValue)
                sectionsToDelete = [sectionIndex.Value];
            else
                throw new ArgumentException(
                    "Either sectionIndex or sectionIndices must be provided for delete operation");

            foreach (var idx in sectionsToDelete)
            {
                if (idx < 0 || idx >= doc.Sections.Count) continue;
                if (doc.Sections.Count <= 1) break; // Don't delete the last section
                doc.Sections.RemoveAt(idx);
            }

            doc.Save(outputPath);
            return
                $"Deleted {sectionsToDelete.Count} section(s). Remaining sections: {doc.Sections.Count}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all sections from the document
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all sections</returns>
    private Task<string> GetSectionsAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);
            var result = new StringBuilder();

            result.AppendLine("=== Document Section Information ===\n");
            result.AppendLine($"Total sections: {doc.Sections.Count}\n");

            if (sectionIndex.HasValue)
            {
                if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                    throw new ArgumentException(
                        $"Section index {sectionIndex.Value} is out of range (document has {doc.Sections.Count} sections)");

                var section = doc.Sections[sectionIndex.Value];
                AppendSectionInfo(result, section, sectionIndex.Value);
            }
            else
            {
                for (var i = 0; i < doc.Sections.Count; i++)
                {
                    var section = doc.Sections[i];
                    AppendSectionInfo(result, section, i);
                    if (i < doc.Sections.Count - 1) result.AppendLine();
                }
            }

            return result.ToString();
        });
    }

    private void AppendSectionInfo(StringBuilder result, Section section, int index)
    {
        result.AppendLine($"[Section {index}]");

        var pageSetup = section.PageSetup;
        result.AppendLine("Page setup:");
        result.AppendLine($"  Paper size: {pageSetup.PaperSize}");
        result.AppendLine($"  Orientation: {pageSetup.Orientation}");
        result.AppendLine($"  Top margin: {pageSetup.TopMargin} points");
        result.AppendLine($"  Bottom margin: {pageSetup.BottomMargin} points");
        result.AppendLine($"  Left margin: {pageSetup.LeftMargin} points");
        result.AppendLine($"  Right margin: {pageSetup.RightMargin} points");
        result.AppendLine($"  Header distance: {pageSetup.HeaderDistance} points");
        result.AppendLine($"  Footer distance: {pageSetup.FooterDistance} points");
        result.AppendLine(
            $"  Page number start: {(pageSetup.RestartPageNumbering ? pageSetup.PageStartingNumber.ToString() : "Inherit from previous section")}");
        result.AppendLine($"  Different first page: {pageSetup.DifferentFirstPageHeaderFooter}");
        result.AppendLine($"  Different odd/even pages: {pageSetup.OddAndEvenPagesHeaderFooter}");
        result.AppendLine($"  Column count: {pageSetup.TextColumns.Count}");

        result.AppendLine();
        result.AppendLine("Content statistics:");
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true);
        var tables = section.Body.GetChildNodes(NodeType.Table, true);
        var shapes = section.Body.GetChildNodes(NodeType.Shape, true);
        result.AppendLine($"  Paragraphs: {paragraphs.Count}");
        result.AppendLine($"  Tables: {tables.Count}");
        result.AppendLine($"  Shapes: {shapes.Count}");

        result.AppendLine();
        result.AppendLine("Headers and footers:");
        var headerCount = 0;
        var footerCount = 0;
        foreach (var hf in section.HeadersFooters.Cast<HeaderFooter>())
            if (hf.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                hf.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                hf.HeaderFooterType == HeaderFooterType.HeaderEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    headerCount++;
            }
            else if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                     hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                     hf.HeaderFooterType == HeaderFooterType.FooterEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    footerCount++;
            }

        result.AppendLine($"  Headers: {headerCount}");
        result.AppendLine($"  Footers: {footerCount}");
    }
}