using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Notes;
using Aspose.Words.Replacing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for footnote and endnote operations in Word documents
/// Merges: WordAddFootnoteTool, WordAddEndnoteTool, WordDeleteFootnoteTool, WordDeleteEndnoteTool,
/// WordEditFootnoteTool, WordEditEndnoteTool, WordGetFootnotesTool, WordGetEndnotesTool
/// </summary>
public class WordNoteTool : IAsposeTool
{
    public string Description => "Manage footnotes and endnotes in Word documents: add, edit, delete, get info";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add_footnote', 'add_endnote', 'delete_footnote', 'delete_endnote', 'edit_footnote', 'edit_endnote', 'get_footnotes', 'get_endnotes'",
                @enum = new[] { "add_footnote", "add_endnote", "delete_footnote", "delete_endnote", "edit_footnote", "edit_endnote", "get_footnotes", "get_endnotes" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

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

    private async Task<string> AddFootnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        var footnoteText = arguments?["noteText"]?.GetValue<string>() ?? throw new ArgumentException("noteText is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>() ?? 0;
        var referenceText = arguments?["referenceText"]?.GetValue<string>();
        var customMark = arguments?["customMark"]?.GetValue<string>();

        var doc = new Document(path);
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var builder = new DocumentBuilder(doc);

        if (!string.IsNullOrEmpty(referenceText))
        {
            var finder = new FindReplaceOptions { MatchCase = false };
            var found = doc.Range.Replace(referenceText, referenceText, finder);
            if (found > 0)
            {
                builder.MoveToDocumentEnd();
                var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
                if (!string.IsNullOrEmpty(customMark))
                {
                    footnote.ReferenceMark = customMark;
                }
            }
            else
            {
                throw new ArgumentException($"Reference text '{referenceText}' not found");
            }
        }
        else if (paragraphIndex.HasValue)
        {
            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");
            }

            var para = paragraphs[paragraphIndex.Value];
            builder.MoveTo(para);
            var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
            if (!string.IsNullOrEmpty(customMark))
            {
                footnote.ReferenceMark = customMark;
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
            var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
            if (!string.IsNullOrEmpty(customMark))
            {
                footnote.ReferenceMark = customMark;
            }
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Footnote added: {outputPath}");
    }

    private async Task<string> AddEndnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        var endnoteText = arguments?["noteText"]?.GetValue<string>() ?? throw new ArgumentException("noteText is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>() ?? 0;
        var referenceText = arguments?["referenceText"]?.GetValue<string>();
        var customMark = arguments?["customMark"]?.GetValue<string>();

        var doc = new Document(path);
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var builder = new DocumentBuilder(doc);

        if (!string.IsNullOrEmpty(referenceText))
        {
            var finder = new FindReplaceOptions { MatchCase = false };
            var found = doc.Range.Replace(referenceText, referenceText, finder);
            if (found > 0)
            {
                builder.MoveToDocumentEnd();
                var endnote = builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
                if (!string.IsNullOrEmpty(customMark))
                {
                    endnote.ReferenceMark = customMark;
                }
            }
            else
            {
                throw new ArgumentException($"Reference text '{referenceText}' not found");
            }
        }
        else if (paragraphIndex.HasValue)
        {
            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");
            }

            var para = paragraphs[paragraphIndex.Value];
            builder.MoveTo(para);
            var endnote = builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
            if (!string.IsNullOrEmpty(customMark))
            {
                endnote.ReferenceMark = customMark;
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
            var endnote = builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
            if (!string.IsNullOrEmpty(customMark))
            {
                endnote.ReferenceMark = customMark;
            }
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Endnote added: {outputPath}");
    }

    private async Task<string> DeleteFootnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var footnoteIndex = arguments?["noteIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Footnote)
            .ToList();

        int deletedCount = 0;

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
        return await Task.FromResult($"Deleted {deletedCount} footnote(s): {outputPath}");
    }

    private async Task<string> DeleteEndnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var endnoteIndex = arguments?["noteIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();

        int deletedCount = 0;

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
        return await Task.FromResult($"Deleted {deletedCount} endnote(s): {outputPath}");
    }

    private async Task<string> EditFootnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var footnoteIndex = arguments?["noteIndex"]?.GetValue<int?>();
        var newText = arguments?["newText"]?.GetValue<string>() ?? throw new ArgumentException("newText is required");

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
            {
                footnote = footnotes[footnoteIndex.Value];
            }
        }
        else if (footnotes.Count > 0)
        {
            footnote = footnotes[0];
        }

        if (footnote == null)
        {
            throw new ArgumentException("Footnote not found");
        }

        footnote.RemoveAllChildren();
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(footnote.FirstParagraph);
        builder.Write(newText);

        doc.Save(outputPath);
        return await Task.FromResult($"Footnote edited: {outputPath}");
    }

    private async Task<string> EditEndnoteAsync(JsonObject? arguments, string path, string outputPath)
    {
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var endnoteIndex = arguments?["noteIndex"]?.GetValue<int?>();
        var newText = arguments?["newText"]?.GetValue<string>() ?? throw new ArgumentException("newText is required");

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
            {
                endnote = endnotes[endnoteIndex.Value];
            }
        }
        else if (endnotes.Count > 0)
        {
            endnote = endnotes[0];
        }

        if (endnote == null)
        {
            throw new ArgumentException("Endnote not found");
        }

        endnote.RemoveAllChildren();
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(endnote.FirstParagraph);
        builder.Write(newText);

        doc.Save(outputPath);
        return await Task.FromResult($"Endnote edited: {outputPath}");
    }

    private async Task<string> GetFootnotesAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        var sb = new StringBuilder();

        sb.AppendLine("=== Footnotes ===");
        sb.AppendLine();

        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Footnote)
            .ToList();

        for (int i = 0; i < footnotes.Count; i++)
        {
            var footnote = footnotes[i];
            sb.AppendLine($"[{i + 1}] Reference Mark: {footnote.ReferenceMark}");
            sb.AppendLine($"    Text: {footnote.ToString(SaveFormat.Text).Trim()}");
            sb.AppendLine($"    Type: {footnote.FootnoteType}");
            sb.AppendLine();
        }

        sb.AppendLine($"Total Footnotes: {footnotes.Count}");

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> GetEndnotesAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        var sb = new StringBuilder();

        sb.AppendLine("=== Endnotes ===");
        sb.AppendLine();

        var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();

        for (int i = 0; i < endnotes.Count; i++)
        {
            var endnote = endnotes[i];
            sb.AppendLine($"[{i + 1}] Reference Mark: {endnote.ReferenceMark}");
            sb.AppendLine($"    Text: {endnote.ToString(SaveFormat.Text).Trim()}");
            sb.AppendLine($"    Type: {endnote.FootnoteType}");
            sb.AppendLine();
        }

        sb.AppendLine($"Total Endnotes: {endnotes.Count}");

        return await Task.FromResult(sb.ToString());
    }
}

