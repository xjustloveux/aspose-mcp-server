using System.Text;
using System.Text.Json;
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
- Add footnote: word_note(operation='add_footnote', path='doc.docx', text='Footnote text', paragraphIndex=0)
- Add endnote: word_note(operation='add_endnote', path='doc.docx', text='Endnote text', paragraphIndex=0)
- Delete footnote: word_note(operation='delete_footnote', path='doc.docx', noteIndex=0)
- Edit footnote: word_note(operation='edit_footnote', path='doc.docx', noteIndex=0, text='Updated footnote')
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
- 'add_footnote': Add a footnote (required params: path, text, paragraphIndex)
- 'add_endnote': Add an endnote (required params: path, text, paragraphIndex)
- 'delete_footnote': Delete a footnote (required params: path, noteIndex)
- 'delete_endnote': Delete an endnote (required params: path, noteIndex)
- 'edit_footnote': Edit a footnote (required params: path, noteIndex, text)
- 'edit_endnote': Edit an endnote (required params: path, noteIndex, text)
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
            text = new
            {
                type = "string",
                description = "Note text content (required for add/edit operations)"
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
            "add_footnote" => await AddNoteAsync(path, outputPath, arguments, FootnoteType.Footnote),
            "add_endnote" => await AddNoteAsync(path, outputPath, arguments, FootnoteType.Endnote),
            "delete_footnote" => await DeleteNoteAsync(path, outputPath, arguments, FootnoteType.Footnote),
            "delete_endnote" => await DeleteNoteAsync(path, outputPath, arguments, FootnoteType.Endnote),
            "edit_footnote" => await EditNoteAsync(path, outputPath, arguments, FootnoteType.Footnote),
            "edit_endnote" => await EditNoteAsync(path, outputPath, arguments, FootnoteType.Endnote),
            "get_footnotes" => await GetNotesAsync(path, FootnoteType.Footnote),
            "get_endnotes" => await GetNotesAsync(path, FootnoteType.Endnote),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets note text from arguments
    /// </summary>
    /// <param name="arguments">JSON arguments containing text parameter</param>
    /// <returns>Note text content</returns>
    private static string GetNoteText(JsonObject? arguments)
    {
        return ArgumentHelper.GetString(arguments, "text");
    }

    /// <summary>
    ///     Gets the display name for a note type
    /// </summary>
    /// <param name="type">Footnote type</param>
    /// <returns>"footnote" or "endnote"</returns>
    private static string GetNoteTypeName(FootnoteType type)
    {
        return type == FootnoteType.Footnote ? "footnote" : "endnote";
    }

    /// <summary>
    ///     Gets all notes of specified type from document
    /// </summary>
    /// <param name="doc">Word document</param>
    /// <param name="type">Footnote type to filter</param>
    /// <returns>List of footnotes/endnotes</returns>
    private static List<Footnote> GetNotes(Document doc, FootnoteType type)
    {
        return doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == type)
            .ToList();
    }

    /// <summary>
    ///     Unified method for adding footnotes and endnotes
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, paragraphIndex, referenceText, customMark</param>
    /// <param name="noteType">Type of note (Footnote or Endnote)</param>
    /// <returns>Success message</returns>
    private Task<string> AddNoteAsync(string path, string outputPath, JsonObject? arguments, FootnoteType noteType)
    {
        return Task.Run(() =>
        {
            var noteText = GetNoteText(arguments);
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var referenceText = ArgumentHelper.GetStringNullable(arguments, "referenceText");
            var customMark = ArgumentHelper.GetStringNullable(arguments, "customMark");
            var typeName = GetNoteTypeName(noteType);

            var doc = new Document(path);
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var builder = new DocumentBuilder(doc);
            Footnote? insertedNote;

            if (!string.IsNullOrEmpty(referenceText))
            {
                // Use IReplacingCallback to insert note at the exact position of referenceText
                var callback = new NoteInsertingCallback(builder, noteType, noteText, customMark);
                var options = new FindReplaceOptions { ReplacingCallback = callback };
                doc.Range.Replace(referenceText, referenceText, options);

                // Check if the callback found and processed the text
                // Note: ReplaceAction.Skip doesn't count as a replacement, so we check InsertedNote instead
                insertedNote = callback.InsertedNote ??
                               throw new ArgumentException($"Reference text '{referenceText}' not found");
            }
            else if (paragraphIndex.HasValue)
            {
                var section = doc.Sections[sectionIndex];

                if (paragraphIndex.Value == -1)
                {
                    // paragraphIndex=-1 means document end
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
                        .ToList();
                    if (bodyParagraphs.Count > 0)
                        builder.MoveTo(bodyParagraphs[^1]);
                    else
                        builder.MoveToDocumentEnd();
                }
                else
                {
                    var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                        throw new ArgumentException(
                            $"paragraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

                    var para = paragraphs[paragraphIndex.Value];

                    // Check for header/footer (endnotes not allowed there)
                    if (noteType == FootnoteType.Endnote)
                    {
                        var parentNode = para.ParentNode;
                        while (parentNode != null)
                        {
                            if (parentNode is HeaderFooter)
                                throw new InvalidOperationException(
                                    $"Endnotes are only allowed inside the main document body. The paragraph at index {paragraphIndex.Value} is located in a header or footer.");
                            if (parentNode is Section || parentNode is Body) break;
                            parentNode = parentNode.ParentNode;
                        }
                    }

                    builder.MoveTo(para);
                }

                insertedNote = builder.InsertFootnote(noteType, noteText);
                if (!string.IsNullOrEmpty(customMark)) insertedNote.ReferenceMark = customMark;
            }
            else
            {
                builder.MoveToDocumentEnd();
                insertedNote = builder.InsertFootnote(noteType, noteText);
                if (!string.IsNullOrEmpty(customMark)) insertedNote.ReferenceMark = customMark;
            }

            doc.Save(outputPath);

            var result = new StringBuilder();
            result.AppendLine($"{char.ToUpper(typeName[0]) + typeName.Substring(1)} added successfully");
            result.AppendLine($"Text: {noteText}");
            if (insertedNote != null && !string.IsNullOrEmpty(insertedNote.ReferenceMark))
                result.AppendLine($"Reference mark: {insertedNote.ReferenceMark}");
            result.AppendLine($"Output: {outputPath}");
            return result.ToString();
        });
    }

    /// <summary>
    ///     Unified method for deleting footnotes and endnotes
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional referenceMark or noteIndex</param>
    /// <param name="noteType">Type of note (Footnote or Endnote)</param>
    /// <returns>Success message with deletion count</returns>
    private Task<string> DeleteNoteAsync(string path, string outputPath, JsonObject? arguments, FootnoteType noteType)
    {
        return Task.Run(() =>
        {
            var referenceMark = ArgumentHelper.GetStringNullable(arguments, "referenceMark");
            var noteIndex = ArgumentHelper.GetIntNullable(arguments, "noteIndex");
            var typeName = GetNoteTypeName(noteType);

            var doc = new Document(path);
            var notes = GetNotes(doc, noteType);

            var deletedCount = 0;

            if (!string.IsNullOrEmpty(referenceMark))
            {
                var note = notes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
                if (note != null)
                {
                    note.Remove();
                    deletedCount = 1;
                }
            }
            else if (noteIndex.HasValue)
            {
                if (noteIndex.Value >= 0 && noteIndex.Value < notes.Count)
                {
                    notes[noteIndex.Value].Remove();
                    deletedCount = 1;
                }
                else
                {
                    throw new ArgumentException(
                        $"Note index {noteIndex.Value} is out of range (document has {notes.Count} {typeName}s, valid index: 0-{notes.Count - 1})");
                }
            }
            else
            {
                foreach (var note in notes)
                {
                    note.Remove();
                    deletedCount++;
                }
            }

            doc.Save(outputPath);
            return $"Deleted {deletedCount} {typeName}(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Unified method for editing footnotes and endnotes
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, optional referenceMark or noteIndex</param>
    /// <param name="noteType">Type of note (Footnote or Endnote)</param>
    /// <returns>Success message with old and new text</returns>
    private Task<string> EditNoteAsync(string path, string outputPath, JsonObject? arguments, FootnoteType noteType)
    {
        return Task.Run(() =>
        {
            var referenceMark = ArgumentHelper.GetStringNullable(arguments, "referenceMark");
            var noteIndex = ArgumentHelper.GetIntNullable(arguments, "noteIndex");
            var newText = GetNoteText(arguments);
            var typeName = GetNoteTypeName(noteType);

            var doc = new Document(path);
            var notes = GetNotes(doc, noteType);

            Footnote? note = null;

            if (!string.IsNullOrEmpty(referenceMark))
            {
                note = notes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
            }
            else if (noteIndex.HasValue)
            {
                if (noteIndex.Value >= 0 && noteIndex.Value < notes.Count)
                    note = notes[noteIndex.Value];
            }
            else if (notes.Count > 0)
            {
                note = notes[0];
            }

            if (note == null)
            {
                var availableInfo = notes.Count > 0
                    ? $" (document has {notes.Count} {typeName}s, valid index: 0-{notes.Count - 1})"
                    : $" (document has no {typeName}s)";
                throw new ArgumentException(
                    $"Specified {typeName} not found{availableInfo}. Use get_{typeName}s operation to view available {typeName}s");
            }

            var oldText = note.ToString(SaveFormat.Text).Trim();
            note.RemoveAllChildren();
            var para = new Paragraph(doc);
            note.AppendChild(para);
            var run = new Run(doc, newText);
            para.AppendChild(run);

            doc.Save(outputPath);

            var result = new StringBuilder();
            result.AppendLine($"{char.ToUpper(typeName[0]) + typeName.Substring(1)} edited successfully");
            result.AppendLine($"Old text: {oldText}");
            result.AppendLine($"New text: {newText}");
            result.AppendLine($"Output: {outputPath}");
            return result.ToString();
        });
    }

    /// <summary>
    ///     Unified method for getting footnotes and endnotes
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="noteType">Type of note (Footnote or Endnote)</param>
    /// <returns>JSON formatted list of notes with indices and content</returns>
    private Task<string> GetNotesAsync(string path, FootnoteType noteType)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var typeName = GetNoteTypeName(noteType);
            var notes = GetNotes(doc, noteType);

            var noteList = new List<object>();
            for (var i = 0; i < notes.Count; i++)
            {
                var note = notes[i];
                noteList.Add(new
                {
                    noteIndex = i,
                    referenceMark = note.ReferenceMark,
                    text = note.ToString(SaveFormat.Text).Trim()
                });
            }

            var result = new
            {
                noteType = typeName,
                count = notes.Count,
                notes = noteList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Callback for inserting notes at the exact position of matched text
    /// </summary>
    private class NoteInsertingCallback(
        DocumentBuilder builder,
        FootnoteType noteType,
        string noteText,
        string? customMark) : IReplacingCallback
    {
        public Footnote? InsertedNote { get; private set; }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Move to the end of the matched node
            var matchNode = args.MatchNode;
            builder.MoveTo(matchNode);

            // Insert the note at this position
            InsertedNote = builder.InsertFootnote(noteType, noteText);
            if (!string.IsNullOrEmpty(customMark))
                InsertedNote.ReferenceMark = customMark;

            // Return Skip to keep the original text and stop after first match
            return ReplaceAction.Skip;
        }
    }
}