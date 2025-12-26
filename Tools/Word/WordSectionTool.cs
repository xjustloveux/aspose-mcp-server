using System.Text.Json;
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
- Get sections: word_section(operation='get', path='doc.docx')

Notes:
- Section break types: NextPage (new page), Continuous (same page), EvenPage, OddPage
- IMPORTANT: Deleting a section will also delete all content within that section (paragraphs, tables, images)
- Use 'get' operation first to see section indices and their content statistics before deleting";

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
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "insert" => await InsertSectionAsync(path, outputPath, arguments),
            "delete" => await DeleteSectionAsync(path, outputPath, arguments),
            "get" => await GetSectionsAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a section break into the document at specified position
    /// </summary>
    /// <param name="path">Source document file path</param>
    /// <param name="outputPath">Output document file path</param>
    /// <param name="arguments">JSON arguments containing sectionBreakType, optional insertAtParagraphIndex, sectionIndex</param>
    /// <returns>Success message with break type and output path</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex or insertAtParagraphIndex is out of range</exception>
    private Task<string> InsertSectionAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sectionBreakType = ArgumentHelper.GetString(arguments, "sectionBreakType");
            var insertAtParagraphIndex = ArgumentHelper.GetIntNullable(arguments, "insertAtParagraphIndex");
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            var breakType = GetSectionStart(sectionBreakType);

            if (insertAtParagraphIndex.HasValue && insertAtParagraphIndex.Value != -1)
            {
                var actualSectionIndex = sectionIndex ?? 0;
                if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
                    throw new ArgumentException(
                        $"sectionIndex must be between 0 and {doc.Sections.Count - 1}, got: {actualSectionIndex}");

                var section = doc.Sections[actualSectionIndex];
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

                if (paragraphs.Count == 0)
                    throw new ArgumentException("Section has no paragraphs to insert section break after");

                if (insertAtParagraphIndex.Value < 0 || insertAtParagraphIndex.Value >= paragraphs.Count)
                    throw new ArgumentException(
                        $"insertAtParagraphIndex must be between 0 and {paragraphs.Count - 1}, got: {insertAtParagraphIndex.Value}");

                builder.MoveTo(paragraphs[insertAtParagraphIndex.Value]);
            }
            else
            {
                builder.MoveToDocumentEnd();
            }

            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.CurrentSection.PageSetup.SectionStart = breakType;

            doc.Save(outputPath);
            return $"Section break inserted ({sectionBreakType}): {outputPath}";
        });
    }

    /// <summary>
    ///     Converts section break type string to SectionStart enum
    /// </summary>
    /// <param name="sectionBreakType">Break type string: NextPage, Continuous, EvenPage, OddPage</param>
    /// <returns>Corresponding SectionStart enum value, defaults to NewPage</returns>
    private static SectionStart GetSectionStart(string sectionBreakType)
    {
        return sectionBreakType switch
        {
            "NextPage" => SectionStart.NewPage,
            "Continuous" => SectionStart.Continuous,
            "EvenPage" => SectionStart.EvenPage,
            "OddPage" => SectionStart.OddPage,
            _ => SectionStart.NewPage
        };
    }

    /// <summary>
    ///     Deletes one or more sections from the document (including all content within)
    /// </summary>
    /// <param name="path">Source document file path</param>
    /// <param name="outputPath">Output document file path</param>
    /// <param name="arguments">JSON arguments containing sectionIndex or sectionIndices array</param>
    /// <returns>Success message with deleted count, remaining sections, and output path</returns>
    /// <exception cref="ArgumentException">Thrown when trying to delete the last section or missing parameters</exception>
    private Task<string> DeleteSectionAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
            var sectionIndicesArray = ArgumentHelper.GetArray(arguments, "sectionIndices", false);

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

            var deletedCount = 0;
            foreach (var idx in sectionsToDelete)
            {
                if (idx < 0 || idx >= doc.Sections.Count) continue;
                if (doc.Sections.Count <= 1) break;
                doc.Sections.RemoveAt(idx);
                deletedCount++;
            }

            doc.Save(outputPath);
            return
                $"Deleted {deletedCount} section(s) with their content. Remaining sections: {doc.Sections.Count}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets section information from the document
    /// </summary>
    /// <param name="path">Source document file path</param>
    /// <param name="arguments">JSON arguments containing optional sectionIndex to get specific section</param>
    /// <returns>JSON formatted string with section details including page setup, content statistics, and headers/footers</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is out of range</exception>
    private Task<string> GetSectionsAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);

            if (sectionIndex.HasValue)
            {
                if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                    throw new ArgumentException(
                        $"sectionIndex must be between 0 and {doc.Sections.Count - 1}, got: {sectionIndex.Value}");

                var section = doc.Sections[sectionIndex.Value];
                var sectionInfo = BuildSectionInfo(section, sectionIndex.Value);

                var result = new
                {
                    totalSections = doc.Sections.Count,
                    section = sectionInfo
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            else
            {
                var sectionList = new List<object>();
                for (var i = 0; i < doc.Sections.Count; i++)
                {
                    var section = doc.Sections[i];
                    sectionList.Add(BuildSectionInfo(section, i));
                }

                var result = new
                {
                    totalSections = doc.Sections.Count,
                    sections = sectionList
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
        });
    }

    /// <summary>
    ///     Builds section information as a structured object
    /// </summary>
    /// <param name="section">Section to get information from</param>
    /// <param name="index">Section index for display</param>
    /// <returns>Anonymous object with section details</returns>
    private static object BuildSectionInfo(Section section, int index)
    {
        var pageSetup = section.PageSetup;

        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true);
        var tables = section.Body.GetChildNodes(NodeType.Table, true);
        var shapes = section.Body.GetChildNodes(NodeType.Shape, true);

        var headerCount = 0;
        var footerCount = 0;
        foreach (var hf in section.HeadersFooters.Cast<HeaderFooter>())
            if (hf.HeaderFooterType is HeaderFooterType.HeaderPrimary or HeaderFooterType.HeaderFirst
                or HeaderFooterType.HeaderEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    headerCount++;
            }
            else if (hf.HeaderFooterType is HeaderFooterType.FooterPrimary or HeaderFooterType.FooterFirst
                     or HeaderFooterType.FooterEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    footerCount++;
            }

        return new
        {
            index,
            sectionBreak = new
            {
                type = GetSectionStartName(pageSetup.SectionStart)
            },
            pageSetup = new
            {
                paperSize = pageSetup.PaperSize.ToString(),
                orientation = pageSetup.Orientation.ToString(),
                margins = new
                {
                    top = pageSetup.TopMargin,
                    bottom = pageSetup.BottomMargin,
                    left = pageSetup.LeftMargin,
                    right = pageSetup.RightMargin
                },
                headerFooterDistance = new
                {
                    header = pageSetup.HeaderDistance,
                    footer = pageSetup.FooterDistance
                },
                pageNumberStart = pageSetup.RestartPageNumbering ? pageSetup.PageStartingNumber : (int?)null,
                differentFirstPage = pageSetup.DifferentFirstPageHeaderFooter,
                differentOddEvenPages = pageSetup.OddAndEvenPagesHeaderFooter,
                columnCount = pageSetup.TextColumns.Count
            },
            contentStatistics = new
            {
                paragraphs = paragraphs.Count,
                tables = tables.Count,
                shapes = shapes.Count
            },
            headersFooters = new
            {
                headers = headerCount,
                footers = footerCount
            }
        };
    }

    /// <summary>
    ///     Converts SectionStart enum to human-readable name
    /// </summary>
    /// <param name="sectionStart">SectionStart enum value</param>
    /// <returns>Human-readable section break type name</returns>
    private static string GetSectionStartName(SectionStart sectionStart)
    {
        return sectionStart switch
        {
            SectionStart.NewPage => "NextPage",
            SectionStart.Continuous => "Continuous",
            SectionStart.EvenPage => "EvenPage",
            SectionStart.OddPage => "OddPage",
            SectionStart.NewColumn => "NewColumn",
            _ => sectionStart.ToString()
        };
    }
}