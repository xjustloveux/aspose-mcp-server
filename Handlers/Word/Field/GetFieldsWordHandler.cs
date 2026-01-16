using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for getting all fields from Word documents.
/// </summary>
public class GetFieldsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_fields";

    /// <summary>
    ///     Gets all fields from the document as JSON with statistics.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: includeCode (default: true), includeResult (default: true)
    /// </param>
    /// <returns>A JSON string containing the list of fields and statistics.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetFieldsParameters(parameters);

        var document = context.Document;
        List<object> fieldsList = [];
        var fieldIndex = 0;

        foreach (var field in document.Range.Fields)
        {
            string? extraInfo = null;
            if (field is FieldHyperlink hyperlinkField)
                extraInfo = $"Address: {hyperlinkField.Address ?? ""}, ScreenTip: {hyperlinkField.ScreenTip ?? ""}";
            else if (field is FieldRef refField)
                extraInfo = $"Bookmark: {refField.BookmarkName ?? ""}";

            fieldsList.Add(new
            {
                index = fieldIndex++,
                type = field.Type.ToString(),
                code = p.IncludeCode ? field.GetFieldCode() : null,
                result = p.IncludeResult ? field.Result ?? "" : null,
                isLocked = field.IsLocked,
                isDirty = field.IsDirty,
                extraInfo
            });
        }

        var statistics = fieldsList
            .GroupBy(f => ((dynamic)f).type as string)
            .OrderBy(g => g.Key)
            .Select(g => new { type = g.Key, count = g.Count() })
            .ToList();

        var result = new
        {
            count = fieldsList.Count,
            fields = fieldsList,
            statisticsByType = statistics
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }

    /// <summary>
    ///     Extracts parameters for the get fields operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetFieldsParameters ExtractGetFieldsParameters(OperationParameters parameters)
    {
        return new GetFieldsParameters(
            parameters.GetOptional("includeCode", true),
            parameters.GetOptional("includeResult", true)
        );
    }

    /// <summary>
    ///     Parameters for the get fields operation.
    /// </summary>
    /// <param name="IncludeCode">Whether to include field code.</param>
    /// <param name="IncludeResult">Whether to include field result.</param>
    private record GetFieldsParameters(bool IncludeCode, bool IncludeResult);
}
