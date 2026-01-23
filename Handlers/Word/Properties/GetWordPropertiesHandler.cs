using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Properties;

namespace AsposeMcpServer.Handlers.Word.Properties;

/// <summary>
///     Handler for getting Word document properties.
/// </summary>
[ResultType(typeof(GetWordPropertiesResult))]
public class GetWordPropertiesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets document properties including built-in and custom properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A JSON string containing built-in properties, statistics, and custom properties.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;

        doc.UpdateWordCount();

        var props = doc.BuiltInDocumentProperties;
        var customProps = doc.CustomDocumentProperties;

        Dictionary<string, CustomPropertyInfo>? customPropsDict = null;
        if (customProps.Count > 0)
        {
            customPropsDict = new Dictionary<string, CustomPropertyInfo>();
            foreach (var prop in customProps)
                customPropsDict[prop.Name] = new CustomPropertyInfo
                {
                    Value = ConvertPropertyValue(prop.Value),
                    Type = prop.Type.ToString()
                };
        }

        return new GetWordPropertiesResult
        {
            BuiltInProperties = new BuiltInPropertiesInfo
            {
                Title = props.Title,
                Subject = props.Subject,
                Author = props.Author,
                Keywords = props.Keywords,
                Comments = props.Comments,
                Category = props.Category,
                Company = props.Company,
                Manager = props.Manager,
                CreatedTime = props.CreatedTime.ToString("O"),
                LastSavedTime = props.LastSavedTime.ToString("O"),
                LastSavedBy = props.LastSavedBy,
                RevisionNumber = props.RevisionNumber
            },
            Statistics = new StatisticsInfo
            {
                WordCount = props.Words,
                CharacterCount = props.Characters,
                PageCount = props.Pages,
                ParagraphCount = props.Paragraphs,
                LineCount = props.Lines
            },
            CustomProperties = customPropsDict
        };
    }

    /// <summary>
    ///     Converts a property value for serialization.
    /// </summary>
    /// <param name="value">The value to convert.</param>
    /// <returns>The converted value.</returns>
    private static object? ConvertPropertyValue(object? value)
    {
        return value switch
        {
            null => null,
            DateTime dt => dt.ToString("O"),
            _ => value
        };
    }
}
