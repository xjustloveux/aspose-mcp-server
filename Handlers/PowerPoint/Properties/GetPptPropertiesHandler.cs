using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.PowerPoint.Properties;

namespace AsposeMcpServer.Handlers.PowerPoint.Properties;

/// <summary>
///     Handler for getting PowerPoint presentation properties.
/// </summary>
[ResultType(typeof(GetPropertiesPptResult))]
public class GetPptPropertiesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets presentation properties.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>JSON string containing the document properties.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        var props = presentation.DocumentProperties;

        var result = new GetPropertiesPptResult
        {
            Title = props.Title,
            Subject = props.Subject,
            Author = props.Author,
            Keywords = props.Keywords,
            Comments = props.Comments,
            Category = props.Category,
            Company = props.Company,
            Manager = props.Manager,
            CreatedTime = props.CreatedTime,
            LastSavedTime = props.LastSavedTime,
            RevisionNumber = props.RevisionNumber
        };

        return result;
    }
}
