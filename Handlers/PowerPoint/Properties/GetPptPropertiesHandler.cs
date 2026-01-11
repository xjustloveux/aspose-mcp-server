using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Properties;

/// <summary>
///     Handler for getting PowerPoint presentation properties.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        var props = presentation.DocumentProperties;

        var result = new
        {
            title = props.Title,
            subject = props.Subject,
            author = props.Author,
            keywords = props.Keywords,
            comments = props.Comments,
            category = props.Category,
            company = props.Company,
            manager = props.Manager,
            createdTime = props.CreatedTime,
            lastSavedTime = props.LastSavedTime,
            revisionNumber = props.RevisionNumber
        };

        return JsonResult(result);
    }
}
