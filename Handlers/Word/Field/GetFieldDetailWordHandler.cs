using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

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
        var p = ExtractGetFieldDetailParameters(parameters);

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (p.FieldIndex < 0 || p.FieldIndex >= fields.Count)
            throw new ArgumentException(
                $"Field index {p.FieldIndex} is out of range (document has {fields.Count} fields)");

        var field = fields[p.FieldIndex];

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
            index = p.FieldIndex,
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

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }

    /// <summary>
    ///     Extracts and validates parameters for the get field detail operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldIndex is not provided.</exception>
    private static GetFieldDetailParameters ExtractGetFieldDetailParameters(OperationParameters parameters)
    {
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");

        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for get_field_detail operation");

        return new GetFieldDetailParameters(fieldIndex.Value);
    }

    /// <summary>
    ///     Parameters for the get field detail operation.
    /// </summary>
    /// <param name="FieldIndex">The index of the field to get details for.</param>
    private record GetFieldDetailParameters(int FieldIndex);
}
