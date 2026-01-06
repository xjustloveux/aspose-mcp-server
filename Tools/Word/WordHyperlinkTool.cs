using System.ComponentModel;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word hyperlinks (add, edit, delete, get)
///     Merges: WordAddHyperlinkTool, WordEditHyperlinkTool, WordDeleteHyperlinkTool, WordGetHyperlinksTool
/// </summary>
[McpServerToolType]
public class WordHyperlinkTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordHyperlinkTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordHyperlinkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word hyperlink operation (add, edit, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Display text for the hyperlink (for add).</param>
    /// <param name="url">URL or target address (for add/edit).</param>
    /// <param name="subAddress">Internal bookmark name for document navigation (for add/edit).</param>
    /// <param name="paragraphIndex">Paragraph index to insert hyperlink after (0-based, for add).</param>
    /// <param name="tooltip">Tooltip text (for add/edit).</param>
    /// <param name="hyperlinkIndex">Hyperlink index (0-based, for edit/delete).</param>
    /// <param name="displayText">New display text (for edit).</param>
    /// <param name="keepText">Keep display text when deleting hyperlink (default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_hyperlink")]
    [Description(@"Manage Word hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: word_hyperlink(operation='add', path='doc.docx', text='Link', url='https://example.com', paragraphIndex=0)
- Edit hyperlink: word_hyperlink(operation='edit', path='doc.docx', hyperlinkIndex=0, url='https://newurl.com')
- Delete hyperlink: word_hyperlink(operation='delete', path='doc.docx', hyperlinkIndex=0)
- Get hyperlinks: word_hyperlink(operation='get', path='doc.docx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a hyperlink (required params: path, text, url)
- 'edit': Edit a hyperlink (required params: path, hyperlinkIndex, url)
- 'delete': Delete a hyperlink (required params: path, hyperlinkIndex)
- 'get': Get all hyperlinks (required params: path)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Display text for the hyperlink (required for add operation)")]
        string? text = null,
        [Description(
            "URL or target address (required for add operation unless subAddress is provided, optional for edit operation)")]
        string? url = null,
        [Description(
            "Internal bookmark name for document navigation (e.g., '_Toc123456'). Use with empty url for internal links. (optional, for add/edit operations)")]
        string? subAddress = null,
        [Description(
            "Paragraph index to insert hyperlink after (0-based, optional, for add operation). When specified, creates a NEW paragraph after the specified paragraph (does not insert into existing paragraph). Valid range: 0 to (total paragraphs - 1), or -1 for document start.")]
        int? paragraphIndex = null,
        [Description("Tooltip text (optional, for add/edit operations)")]
        string? tooltip = null,
        [Description("Hyperlink index (0-based, required for edit/delete operations)")]
        int? hyperlinkIndex = null,
        [Description("New display text (optional, for edit operation)")]
        string? displayText = null,
        [Description(
            "Keep display text when deleting hyperlink (unlink instead of remove, optional, default: false, for delete operation)")]
        bool keepText = false)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddHyperlink(ctx, outputPath, text!, url, subAddress, paragraphIndex, tooltip),
            "edit" => EditHyperlink(ctx, outputPath, hyperlinkIndex ?? 0, url, subAddress, displayText, tooltip),
            "delete" => DeleteHyperlink(ctx, outputPath, hyperlinkIndex ?? 0, keepText),
            "get" => GetHyperlinks(ctx),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The display text for the hyperlink.</param>
    /// <param name="url">The URL or target address.</param>
    /// <param name="subAddress">The internal bookmark name for document navigation.</param>
    /// <param name="paragraphIndex">The paragraph index to insert after.</param>
    /// <param name="tooltip">The tooltip text.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither url nor subAddress is provided, or URL format is invalid, or
    ///     paragraph index is out of range.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph cannot be found or accessed.</exception>
    private static string AddHyperlink(DocumentContext<Document> ctx, string? outputPath, string text, string? url,
        string? subAddress, int? paragraphIndex, string? tooltip)
    {
        // Validate: either url or subAddress must be provided
        if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(subAddress))
            throw new ArgumentException("Either 'url' or 'subAddress' must be provided for add operation");

        // Validate URL format if provided
        if (!string.IsNullOrEmpty(url))
            ValidateUrlFormat(url);

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                // Insert at the beginning - create new paragraph
                if (paragraphs.Count > 0)
                {
                    if (paragraphs[0] is Paragraph firstPara)
                    {
                        // Insert new paragraph before the first paragraph
                        var newPara = new Paragraph(doc);
                        doc.FirstSection.Body.InsertBefore(newPara, firstPara);
                        builder.MoveTo(newPara);
                    }
                    else
                    {
                        builder.MoveToDocumentStart();
                    }
                }
                else
                {
                    builder.MoveToDocumentStart();
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                // Insert after the specified paragraph - create new paragraph
                if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                {
                    // Insert new paragraph after the target paragraph
                    var newPara = new Paragraph(doc);
                    var parentNode = targetPara.ParentNode;
                    if (parentNode != null)
                    {
                        parentNode.InsertAfter(newPara, targetPara);
                        builder.MoveTo(newPara);
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Unable to find parent node of paragraph at index {paragraphIndex.Value}");
                    }
                }
                else
                {
                    throw new InvalidOperationException(
                        $"Unable to find paragraph at index {paragraphIndex.Value}");
                }
            }
            else
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
        }
        else
        {
            // Default: Move to end of document
            builder.MoveToDocumentEnd();
        }

        if (!string.IsNullOrEmpty(subAddress))
            builder.InsertHyperlink(text, subAddress, true);
        else
            builder.InsertHyperlink(text, url!, false);

        var fields = doc.Range.Fields;
        if (fields.Count > 0)
        {
            var lastField = fields[^1];
            if (lastField is FieldHyperlink hyperlinkField)
            {
                if (!string.IsNullOrEmpty(tooltip))
                    hyperlinkField.ScreenTip = tooltip;
                // Set both Address and SubAddress for combined links
                if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(subAddress))
                {
                    hyperlinkField.Address = url;
                    hyperlinkField.SubAddress = subAddress;
                }
            }
        }

        ctx.Save(outputPath);

        var result = "Hyperlink added successfully\n";
        result += $"Display text: {text}\n";
        if (!string.IsNullOrEmpty(url)) result += $"URL: {url}\n";
        if (!string.IsNullOrEmpty(subAddress)) result += $"SubAddress (bookmark): {subAddress}\n";
        if (!string.IsNullOrEmpty(tooltip)) result += $"Tooltip: {tooltip}\n";
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
                result += "Insert position: beginning of document\n";
            else
                result += $"Insert position: after paragraph #{paragraphIndex.Value}\n";
        }
        else
        {
            result += "Insert position: end of document\n";
        }

        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="hyperlinkIndex">The zero-based hyperlink index.</param>
    /// <param name="url">The new URL or target address.</param>
    /// <param name="subAddress">The new internal bookmark name.</param>
    /// <param name="displayText">The new display text.</param>
    /// <param name="tooltip">The new tooltip text.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the hyperlink index is out of range or URL format is invalid.</exception>
    private static string EditHyperlink(DocumentContext<Document> ctx, string? outputPath, int hyperlinkIndex,
        string? url, string? subAddress, string? displayText, string? tooltip)
    {
        var doc = ctx.Document;
        var hyperlinkFields = GetAllHyperlinks(doc);

        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinkFields.Count)
        {
            var availableInfo = hyperlinkFields.Count > 0
                ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                : " (document has no hyperlinks)";
            throw new ArgumentException(
                $"Hyperlink index {hyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
        }

        var hyperlinkField = hyperlinkFields[hyperlinkIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(url))
        {
            ValidateUrlFormat(url);
            hyperlinkField.Address = url;
            changes.Add($"URL: {url}");
        }

        if (!string.IsNullOrEmpty(subAddress))
        {
            hyperlinkField.SubAddress = subAddress;
            changes.Add($"SubAddress: {subAddress}");
        }

        if (!string.IsNullOrEmpty(displayText))
        {
            hyperlinkField.Result = displayText;
            changes.Add($"Display text: {displayText}");
        }

        if (!string.IsNullOrEmpty(tooltip))
        {
            hyperlinkField.ScreenTip = tooltip;
            changes.Add($"Tooltip: {tooltip}");
        }

        hyperlinkField.Update();

        ctx.Save(outputPath);

        var result = $"Hyperlink #{hyperlinkIndex} edited successfully\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}\n";
        else
            result += "No change parameters provided\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Deletes a hyperlink from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="hyperlinkIndex">The zero-based hyperlink index.</param>
    /// <param name="keepText">Whether to keep the display text when deleting.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the hyperlink index is out of range.</exception>
    private static string DeleteHyperlink(DocumentContext<Document> ctx, string? outputPath, int hyperlinkIndex,
        bool keepText)
    {
        var doc = ctx.Document;
        var hyperlinkFields = GetAllHyperlinks(doc);

        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinkFields.Count)
        {
            var availableInfo = hyperlinkFields.Count > 0
                ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                : " (document has no hyperlinks)";
            throw new ArgumentException(
                $"Hyperlink index {hyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
        }

        var hyperlinkField = hyperlinkFields[hyperlinkIndex];
        var displayTextValue = hyperlinkField.Result ?? "";
        var address = hyperlinkField.Address ?? "";

        if (keepText)
            hyperlinkField.Unlink();
        else
            hyperlinkField.Remove();

        ctx.Save(outputPath);

        var remainingCount = GetAllHyperlinks(doc).Count;

        var result = $"Hyperlink #{hyperlinkIndex} deleted successfully\n";
        result += $"Display text: {displayTextValue}\n";
        result += $"Address: {address}\n";
        result += $"Keep text: {(keepText ? "Yes (unlinked)" : "No (removed)")}\n";
        result += $"Remaining hyperlinks in document: {remainingCount}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Gets all hyperlink fields from the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <returns>A list of FieldHyperlink objects representing all hyperlinks.</returns>
    private static List<FieldHyperlink> GetAllHyperlinks(Document doc)
    {
        return doc.Range.Fields.OfType<FieldHyperlink>().ToList();
    }

    /// <summary>
    ///     Validates URL format to prevent invalid field commands.
    /// </summary>
    /// <param name="url">The URL to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the URL format is invalid.</exception>
    private static void ValidateUrlFormat(string url)
    {
        var validPrefixes = new[] { "http://", "https://", "mailto:", "ftp://", "file://", "#" };
        if (!validPrefixes.Any(prefix => url.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException(
                $"Invalid URL format: '{url}'. URL must start with http://, https://, mailto:, ftp://, file://, or # (for internal links)");
    }

    /// <summary>
    ///     Gets all hyperlinks from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing hyperlink information.</returns>
    private static string GetHyperlinks(DocumentContext<Document> ctx)
    {
        var doc = ctx.Document;
        var hyperlinkFields = GetAllHyperlinks(doc);

        if (hyperlinkFields.Count == 0)
            return JsonSerializer.Serialize(new
                { count = 0, hyperlinks = Array.Empty<object>(), message = "No hyperlinks found in document" });

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        List<object> hyperlinkList = [];
        for (var index = 0; index < hyperlinkFields.Count; index++)
        {
            var hyperlinkField = hyperlinkFields[index];
            var displayText = "";
            var address = "";
            var subAddress = "";
            var tooltip = "";
            int? paragraphIndexValue = null;

            try
            {
                displayText = hyperlinkField.Result ?? "";
                address = hyperlinkField.Address ?? "";
                subAddress = hyperlinkField.SubAddress ?? "";
                tooltip = hyperlinkField.ScreenTip ?? "";

                var fieldStart = hyperlinkField.Start;
                if (fieldStart?.ParentNode is Paragraph para)
                {
                    paragraphIndexValue = paragraphs.IndexOf(para);
                    if (paragraphIndexValue == -1) paragraphIndexValue = null;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Error reading hyperlink properties: {ex.Message}");
            }

            hyperlinkList.Add(new
            {
                index,
                displayText,
                address,
                subAddress = string.IsNullOrEmpty(subAddress) ? null : subAddress,
                tooltip = string.IsNullOrEmpty(tooltip) ? null : tooltip,
                paragraphIndex = paragraphIndexValue
            });
        }

        var result = new
        {
            count = hyperlinkFields.Count,
            hyperlinks = hyperlinkList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}