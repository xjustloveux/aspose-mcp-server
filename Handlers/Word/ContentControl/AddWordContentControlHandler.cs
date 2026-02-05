using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.ContentControl;

/// <summary>
///     Handler for adding content controls (structured document tags) to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddWordContentControlHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a content control to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: type (PlainText, RichText, DropDownList, DatePicker, CheckBox, Picture, ComboBox)
    ///     Optional: tag, title, value, items (comma-separated for DropDownList/ComboBox),
    ///     lockContents (default: false), lockDeletion (default: false), paragraphIndex
    /// </param>
    /// <returns>Success message with content control details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        var sdtType = ResolveSdtType(p.Type);

        MoveToInsertPosition(builder, doc, p.ParagraphIndex);

        var sdt = new StructuredDocumentTag(doc, sdtType, MarkupLevel.Inline);

        if (!string.IsNullOrEmpty(p.Tag)) sdt.Tag = p.Tag;
        if (!string.IsNullOrEmpty(p.Title)) sdt.Title = p.Title;
        sdt.LockContents = p.LockContents;
        sdt.LockContentControl = p.LockDeletion;

        ConfigureContentControl(sdt, sdtType, p);

        builder.InsertNode(sdt);

        MarkModified(context);

        return new SuccessResult
        {
            Message = BuildMessage(p)
        };
    }

    /// <summary>
    ///     Resolves the SdtType from a type string.
    /// </summary>
    /// <param name="type">The type string.</param>
    /// <returns>The corresponding SdtType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the type is unknown.</exception>
    private static SdtType ResolveSdtType(string type)
    {
        return type.ToLowerInvariant() switch
        {
            "plaintext" => SdtType.PlainText,
            "richtext" => SdtType.RichText,
            "dropdownlist" => SdtType.DropDownList,
            "datepicker" => SdtType.Date,
            "checkbox" => SdtType.Checkbox,
            "picture" => SdtType.Picture,
            "combobox" => SdtType.ComboBox,
            _ => throw new ArgumentException(
                $"Unknown content control type: {type}. Supported: PlainText, RichText, DropDownList, DatePicker, CheckBox, Picture, ComboBox")
        };
    }

    /// <summary>
    ///     Configures the content control based on its type and parameters.
    /// </summary>
    /// <param name="sdt">The structured document tag to configure.</param>
    /// <param name="sdtType">The type of the content control.</param>
    /// <param name="p">The add parameters.</param>
    private static void ConfigureContentControl(StructuredDocumentTag sdt, SdtType sdtType, AddParameters p)
    {
        if (sdtType is SdtType.DropDownList or SdtType.ComboBox && !string.IsNullOrEmpty(p.Items))
        {
            var items = p.Items.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var item in items)
                sdt.ListItems.Add(new SdtListItem(item, item));
        }

        if (!string.IsNullOrEmpty(p.Value))
        {
            if (sdtType == SdtType.Checkbox)
            {
                sdt.Checked = p.Value.Equals("true", StringComparison.OrdinalIgnoreCase);
            }
            else if (sdtType is SdtType.PlainText or SdtType.RichText)
            {
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(sdt.Document, p.Value));
            }
        }
    }

    /// <summary>
    ///     Moves the document builder to the specified insert position.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="paragraphIndex">The paragraph index to move to, or null for end of document.</param>
    /// <exception cref="ArgumentException">Thrown when the paragraph index is out of range.</exception>
    private static void MoveToInsertPosition(DocumentBuilder builder, Document doc, int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
        {
            builder.MoveToDocumentEnd();
            return;
        }

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphs.Count == 0)
        {
            builder.MoveToDocumentEnd();
            return;
        }

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        if (paragraphs[paragraphIndex.Value] is Aspose.Words.Paragraph para)
            builder.MoveTo(para);
        else
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");
    }

    /// <summary>
    ///     Builds the result message for a successful content control addition.
    /// </summary>
    /// <param name="p">The add parameters.</param>
    /// <returns>A formatted result message.</returns>
    private static string BuildMessage(AddParameters p)
    {
        var message = $"Content control of type '{p.Type}' added successfully.";
        if (!string.IsNullOrEmpty(p.Tag)) message += $" Tag: {p.Tag}.";
        if (!string.IsNullOrEmpty(p.Title)) message += $" Title: {p.Title}.";
        return message;
    }

    /// <summary>
    ///     Extracts and validates parameters for the add content control operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when the type parameter is missing.</exception>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        var type = parameters.GetOptional<string?>("type");
        if (string.IsNullOrEmpty(type))
            throw new ArgumentException("Content control type is required for add operation");

        return new AddParameters(
            type,
            parameters.GetOptional<string?>("tag"),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("value"),
            parameters.GetOptional<string?>("items"),
            parameters.GetOptional("lockContents", false),
            parameters.GetOptional("lockDeletion", false),
            parameters.GetOptional<int?>("paragraphIndex")
        );
    }

    /// <summary>
    ///     Parameters for the add content control operation.
    /// </summary>
    /// <param name="Type">The content control type.</param>
    /// <param name="Tag">The tag identifier.</param>
    /// <param name="Title">The display title.</param>
    /// <param name="Value">The initial value.</param>
    /// <param name="Items">Comma-separated items for DropDownList/ComboBox.</param>
    /// <param name="LockContents">Whether to lock contents.</param>
    /// <param name="LockDeletion">Whether to lock deletion.</param>
    /// <param name="ParagraphIndex">The paragraph index to insert at.</param>
    private sealed record AddParameters(
        string Type,
        string? Tag,
        string? Title,
        string? Value,
        string? Items,
        bool LockContents,
        bool LockDeletion,
        int? ParagraphIndex);
}
