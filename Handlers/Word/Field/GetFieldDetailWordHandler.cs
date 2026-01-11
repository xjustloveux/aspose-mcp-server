using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for getting detailed information about a specific field in Word documents.
/// </summary>
public class GetFieldDetailWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_field_detail";

    /// <summary>
    ///     Gets detailed information about a specific field.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldIndex
    /// </param>
    /// <returns>A JSON string containing the field details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");

        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for get_field_detail operation");

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            throw new ArgumentException(
                $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

        var field = fields[fieldIndex.Value];

        string? address = null, screenTip = null, bookmarkName = null;
        if (field is FieldHyperlink hyperlinkField)
        {
            address = hyperlinkField.Address;
            screenTip = hyperlinkField.ScreenTip;
        }
        else if (field is FieldRef refField)
        {
            bookmarkName = refField.BookmarkName;
        }

        var result = new
        {
            index = fieldIndex.Value,
            type = field.Type.ToString(),
            typeCode = (int)field.Type,
            code = field.GetFieldCode(),
            result = field.Result,
            isLocked = field.IsLocked,
            isDirty = field.IsDirty,
            hyperlinkAddress = address,
            hyperlinkScreenTip = screenTip,
            bookmarkName
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}
