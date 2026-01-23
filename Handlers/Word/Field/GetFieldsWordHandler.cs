using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Field;
using WordFieldInfo = AsposeMcpServer.Results.Word.Field.FieldInfo;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for getting all fields from Word documents.
/// </summary>
[ResultType(typeof(GetFieldsWordResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetFieldsParameters(parameters);

        var document = context.Document;
        List<WordFieldInfo> fieldsList = [];
        var fieldIndex = 0;

        foreach (var field in document.Range.Fields)
        {
            string? extraInfo = null;
            if (field is FieldHyperlink hyperlinkField)
                extraInfo = $"Address: {hyperlinkField.Address ?? ""}, ScreenTip: {hyperlinkField.ScreenTip ?? ""}";
            else if (field is FieldRef refField)
                extraInfo = $"Bookmark: {refField.BookmarkName ?? ""}";

            fieldsList.Add(new WordFieldInfo
            {
                Index = fieldIndex++,
                Type = field.Type.ToString(),
                Code = p.IncludeCode ? field.GetFieldCode() : null,
                Result = p.IncludeResult ? field.Result ?? "" : null,
                IsLocked = field.IsLocked,
                IsDirty = field.IsDirty,
                ExtraInfo = extraInfo
            });
        }

        var statistics = fieldsList
            .GroupBy(f => f.Type)
            .OrderBy(g => g.Key)
            .Select(g => new FieldTypeStatistics { Type = g.Key, Count = g.Count() })
            .ToList();

        return new GetFieldsWordResult
        {
            Count = fieldsList.Count,
            Fields = fieldsList,
            StatisticsByType = statistics
        };
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
    private sealed record GetFieldsParameters(bool IncludeCode, bool IncludeResult);
}
