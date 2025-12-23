using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Replacing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for footnote and endnote operations in Word documents
///     Merges: WordAddFootnoteTool, WordAddEndnoteTool, WordDeleteFootnoteTool, WordDeleteEndnoteTool,
///     WordEditFootnoteTool, WordEditEndnoteTool, WordGetFootnotesTool, WordGetEndnotesTool
/// </summary>
public class WordNoteTool : IAsposeTool
{
    public string Description =>
        @"Manage footnotes and endnotes in Word documents. Supports 8 operations: add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote, edit_endnote, get_footnotes, get_endnotes.

Usage examples:
- Add footnote: word_note(operation='add_footnote', path='doc.docx', noteText='Footnote text', paragraphIndex=0, runIndex=0)
- Add endnote: word_note(operation='add_endnote', path='doc.docx', noteText='Endnote text', paragraphIndex=0)
- Delete footnote: word_note(operation='delete_footnote', path='doc.docx', noteIndex=0)
- Edit footnote: word_note(operation='edit_footnote', path='doc.docx', noteIndex=0, newText='Updated footnote')
- Get footnotes: word_note(operation='get_footnotes', path='doc.docx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add_footnote': Add a footnote (required params: path, noteText, paragraphIndex)
- 'add_endnote': Add an endnote (required params: path, noteText, paragraphIndex)
- 'delete_footnote': Delete a footnote (required params: path, noteIndex)
- 'delete_endnote': Delete an endnote (required params: path, noteIndex)
- 'edit_footnote': Edit a footnote (required params: path, noteIndex, newText)
- 'edit_endnote': Edit an endnote (required params: path, noteIndex, newText)
- 'get_footnotes': Get all footnotes (required params: path)
- 'get_endnotes': Get all endnotes (required params: path)",
                @enum = new[]
                {
                    "add_footnote", "add_endnote", "delete_footnote", "delete_endnote", "edit_footnote", "edit_endnote",
                    "get_footnotes", "get_endnotes"
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
            // Common parameters
            noteText = new
            {
                type = "string",
                description = "Note text (required for add operations, newText for edit operations)"
            },
            newText = new
            {
                type = "string",
                description = "New note text (required for edit operations)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, optional)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            },
            referenceText = new
            {
                type = "string",
                description = "Reference text in document (optional, if not provided inserts at paragraph end)"
            },
            customMark = new
            {
                type = "string",
                description = "Custom note mark (optional, if not provided uses auto-numbering)"
            },
            referenceMark = new
            {
                type = "string",
                description = "Reference mark of note to delete/edit (optional)"
            },
            noteIndex = new
            {
                type = "number",
                description = "Note index (0-based, optional, for delete/edit operations)"
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

        return operation switch
        {
            "add_footnote" => await AddFootnoteAsync(arguments, path, outputPath),
            "add_endnote" => await AddEndnoteAsync(arguments, path, outputPath),
            "delete_footnote" => await DeleteFootnoteAsync(arguments, path, outputPath),
            "delete_endnote" => await DeleteEndnoteAsync(arguments, path, outputPath),
            "edit_footnote" => await EditFootnoteAsync(arguments, path, outputPath),
            "edit_endnote" => await EditEndnoteAsync(arguments, path, outputPath),
            "get_footnotes" => await GetFootnotesAsync(arguments, path),
            "get_endnotes" => await GetEndnotesAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a footnote to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing noteText, optional paragraphIndex, sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> AddFootnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var footnoteText = ArgumentHelper.GetString(arguments, "noteText");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var referenceText = ArgumentHelper.GetStringNullable(arguments, "referenceText");
            var customMark = ArgumentHelper.GetStringNullable(arguments, "customMark");

            var doc = new Document(path);
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var builder = new DocumentBuilder(doc);

            if (!string.IsNullOrEmpty(referenceText))
            {
                var finder = new FindReplaceOptions { MatchCase = false };
                var found = doc.Range.Replace(referenceText, referenceText, finder);
                if (found > 0)
                {
                    builder.MoveToDocumentEnd();
                    var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
                    if (!string.IsNullOrEmpty(customMark)) footnote.ReferenceMark = customMark;
                }
                else
                {
                    throw new ArgumentException($"Reference text '{referenceText}' not found");
                }
            }
            else if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value == -1)
                {
                    // paragraphIndex=-1 means document end - move to last paragraph in Body
                    var section = doc.Sections[sectionIndex];
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
                        .ToList();
                    if (bodyParagraphs.Count > 0)
                    {
                        var lastPara = bodyParagraphs[^1];
                        builder.MoveTo(lastPara);
                    }
                    else
                    {
                        // No paragraphs in body, move to document end
                        builder.MoveToDocumentEnd();
                    }

                    var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
                    if (!string.IsNullOrEmpty(customMark)) footnote.ReferenceMark = customMark;
                }
                else
                {
                    var section = doc.Sections[sectionIndex];
                    var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                        throw new ArgumentException(
                            $"paragraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

                    var para = paragraphs[paragraphIndex.Value];
                    builder.MoveTo(para);
                    var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
                    if (!string.IsNullOrEmpty(customMark)) footnote.ReferenceMark = customMark;
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
                var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
                if (!string.IsNullOrEmpty(customMark)) footnote.ReferenceMark = customMark;
            }

            doc.Save(outputPath);
            return $"Footnote added: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds an endnote to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing noteText, optional paragraphIndex, sectionIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> AddEndnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var endnoteText = ArgumentHelper.GetString(arguments, "noteText");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var referenceText = ArgumentHelper.GetStringNullable(arguments, "referenceText");
            var customMark = ArgumentHelper.GetStringNullable(arguments, "customMark");

            var doc = new Document(path);
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var builder = new DocumentBuilder(doc);

            if (!string.IsNullOrEmpty(referenceText))
            {
                var finder = new FindReplaceOptions { MatchCase = false };
                var found = doc.Range.Replace(referenceText, referenceText, finder);
                if (found > 0)
                {
                    builder.MoveToDocumentEnd();
                    var endnote = builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
                    if (!string.IsNullOrEmpty(customMark)) endnote.ReferenceMark = customMark;
                }
                else
                {
                    throw new ArgumentException($"Reference text '{referenceText}' not found");
                }
            }
            else if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value == -1)
                {
                    // paragraphIndex=-1 means document end - move to last paragraph in Body
                    var section = doc.Sections[sectionIndex];
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
                        .ToList();
                    if (bodyParagraphs.Count > 0)
                    {
                        var lastPara = bodyParagraphs[^1];
                        builder.MoveTo(lastPara);
                    }
                    else
                    {
                        // No paragraphs in body, move to document end
                        builder.MoveToDocumentEnd();
                    }

                    var endnote = builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
                    if (!string.IsNullOrEmpty(customMark)) endnote.ReferenceMark = customMark;
                }
                else
                {
                    var section = doc.Sections[sectionIndex];
                    var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                        throw new ArgumentException(
                            $"paragraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

                    var para = paragraphs[paragraphIndex.Value];

                    var parentNode = para.ParentNode;
                    while (parentNode != null)
                    {
                        if (parentNode is HeaderFooter)
                            throw new InvalidOperationException(
                                $"Endnotes are only allowed inside the main document body. The paragraph at index {paragraphIndex.Value} is located in a header or footer. Please use a paragraph index that refers to a paragraph in the main document body.");
                        if (parentNode is Section || parentNode is Body) break; // We're in the main body
                        parentNode = parentNode.ParentNode;
                    }

                    builder.MoveTo(para);
                    try
                    {
                        var endnote = builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
                        if (!string.IsNullOrEmpty(customMark)) endnote.ReferenceMark = customMark;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            $"Failed to insert endnote: {ex.Message}. Endnotes can only be inserted in the main document body, not in headers, footers, or other special sections.",
                            ex);
                    }
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
                var endnote = builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
                if (!string.IsNullOrEmpty(customMark)) endnote.ReferenceMark = customMark;
            }

            doc.Save(outputPath);
            return $"Endnote added: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a footnote from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing noteIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteFootnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var referenceMark = ArgumentHelper.GetStringNullable(arguments, "referenceMark");
            var footnoteIndex = ArgumentHelper.GetIntNullable(arguments, "noteIndex");

            var doc = new Document(path);
            var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
                .Where(f => f.FootnoteType == FootnoteType.Footnote)
                .ToList();

            var deletedCount = 0;

            if (!string.IsNullOrEmpty(referenceMark))
            {
                var footnote = footnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
                if (footnote != null)
                {
                    footnote.Remove();
                    deletedCount = 1;
                }
            }
            else if (footnoteIndex.HasValue)
            {
                if (footnoteIndex.Value >= 0 && footnoteIndex.Value < footnotes.Count)
                {
                    footnotes[footnoteIndex.Value].Remove();
                    deletedCount = 1;
                }
            }
            else
            {
                foreach (var footnote in footnotes)
                {
                    footnote.Remove();
                    deletedCount++;
                }
            }

            doc.Save(outputPath);
            return $"Deleted {deletedCount} footnote(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes an endnote from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing noteIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteEndnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var referenceMark = ArgumentHelper.GetStringNullable(arguments, "referenceMark");
            var endnoteIndex = ArgumentHelper.GetIntNullable(arguments, "noteIndex");

            var doc = new Document(path);
            var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
                .Where(f => f.FootnoteType == FootnoteType.Endnote)
                .ToList();

            var deletedCount = 0;

            if (!string.IsNullOrEmpty(referenceMark))
            {
                var endnote = endnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
                if (endnote != null)
                {
                    endnote.Remove();
                    deletedCount = 1;
                }
            }
            else if (endnoteIndex.HasValue)
            {
                if (endnoteIndex.Value >= 0 && endnoteIndex.Value < endnotes.Count)
                {
                    endnotes[endnoteIndex.Value].Remove();
                    deletedCount = 1;
                }
            }
            else
            {
                foreach (var endnote in endnotes)
                {
                    endnote.Remove();
                    deletedCount++;
                }
            }

            doc.Save(outputPath);
            return $"Deleted {deletedCount} endnote(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Edits a footnote
    /// </summary>
    /// <param name="arguments">JSON arguments containing noteIndex, noteText</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> EditFootnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var referenceMark = ArgumentHelper.GetStringNullable(arguments, "referenceMark");
            var footnoteIndex = ArgumentHelper.GetIntNullable(arguments, "noteIndex");
            var newText = ArgumentHelper.GetString(arguments, "newText");

            var doc = new Document(path);
            var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
                .Where(f => f.FootnoteType == FootnoteType.Footnote)
                .ToList();

            Footnote? footnote = null;

            if (!string.IsNullOrEmpty(referenceMark))
            {
                footnote = footnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
            }
            else if (footnoteIndex.HasValue)
            {
                if (footnoteIndex.Value >= 0 && footnoteIndex.Value < footnotes.Count)
                    footnote = footnotes[footnoteIndex.Value];
            }
            else if (footnotes.Count > 0)
            {
                footnote = footnotes[0];
            }

            if (footnote == null)
            {
                var availableInfo = footnotes.Count > 0
                    ? $" (document has {footnotes.Count} footnotes, valid index: 0-{footnotes.Count - 1})"
                    : " (document has no footnotes)";
                throw new ArgumentException(
                    $"Specified footnote not found{availableInfo}. Use get_footnotes operation to view available footnotes");
            }

            footnote.RemoveAllChildren();
            var builder = new DocumentBuilder(doc);
            builder.MoveTo(footnote.FirstParagraph);
            builder.Write(newText);

            doc.Save(outputPath);
            return $"Footnote edited: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an endnote
    /// </summary>
    /// <param name="arguments">JSON arguments containing noteIndex, noteText</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> EditEndnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var referenceMark = ArgumentHelper.GetStringNullable(arguments, "referenceMark");
            var endnoteIndex = ArgumentHelper.GetIntNullable(arguments, "noteIndex");
            var newText = ArgumentHelper.GetString(arguments, "newText");

            var doc = new Document(path);
            var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
                .Where(f => f.FootnoteType == FootnoteType.Endnote)
                .ToList();

            Footnote? endnote = null;

            if (!string.IsNullOrEmpty(referenceMark))
            {
                endnote = endnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
            }
            else if (endnoteIndex.HasValue)
            {
                if (endnoteIndex.Value >= 0 && endnoteIndex.Value < endnotes.Count)
                    endnote = endnotes[endnoteIndex.Value];
            }
            else if (endnotes.Count > 0)
            {
                endnote = endnotes[0];
            }

            if (endnote == null)
            {
                var availableInfo = endnotes.Count > 0
                    ? $" (document has {endnotes.Count} endnotes, valid index: 0-{endnotes.Count - 1})"
                    : " (document has no endnotes)";
                throw new ArgumentException(
                    $"Specified endnote not found{availableInfo}. Use get_endnotes operation to view available endnotes");
            }

            endnote.RemoveAllChildren();
            var builder = new DocumentBuilder(doc);
            builder.MoveTo(endnote.FirstParagraph);
            builder.Write(newText);

            doc.Save(outputPath);
            return $"Endnote edited: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all footnotes from the document
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all footnotes</returns>
    private Task<string> GetFootnotesAsync(JsonObject? _, string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var sb = new StringBuilder();

            sb.AppendLine("=== Footnotes ===");
            sb.AppendLine();

            var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
                .Where(f => f.FootnoteType == FootnoteType.Footnote)
                .ToList();

            for (var i = 0; i < footnotes.Count; i++)
            {
                var footnote = footnotes[i];
                sb.AppendLine($"[{i + 1}] Reference Mark: {footnote.ReferenceMark}");
                sb.AppendLine($"    Text: {footnote.ToString(SaveFormat.Text).Trim()}");
                sb.AppendLine($"    Type: {footnote.FootnoteType}");
                sb.AppendLine();
            }

            sb.AppendLine($"Total Footnotes: {footnotes.Count}");

            return sb.ToString();
        });
    }

    /// <summary>
    ///     Gets all endnotes from the document
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all endnotes</returns>
    private Task<string> GetEndnotesAsync(JsonObject? _, string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var sb = new StringBuilder();

            sb.AppendLine("=== Endnotes ===");
            sb.AppendLine();

            var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
                .Where(f => f.FootnoteType == FootnoteType.Endnote)
                .ToList();

            for (var i = 0; i < endnotes.Count; i++)
            {
                var endnote = endnotes[i];
                sb.AppendLine($"[{i + 1}] Reference Mark: {endnote.ReferenceMark}");
                sb.AppendLine($"    Text: {endnote.ToString(SaveFormat.Text).Trim()}");
                sb.AppendLine($"    Type: {endnote.FootnoteType}");
                sb.AppendLine();
            }

            sb.AppendLine($"Total Endnotes: {endnotes.Count}");

            return sb.ToString();
        });
    }
}