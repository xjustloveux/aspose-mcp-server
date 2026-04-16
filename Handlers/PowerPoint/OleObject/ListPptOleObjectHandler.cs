using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.PowerPoint.OleObject;

/// <summary>
///     Handler for the <c>list</c> operation on <c>ppt_ole_object</c>.
/// </summary>
[ResultType(typeof(OleListResult))]
public sealed class ListPptOleObjectHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "list";

    /// <summary>Executes the <c>list</c> operation across all slides.</summary>
    /// <param name="context">Operation context; <c>Presentation</c> must be non-null.</param>
    /// <param name="parameters">Unused.</param>
    /// <returns>An <see cref="OleListResult" /> enumerating every OLE frame.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="context" /> is null.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        var presentation = context.Document;
        var items = new List<OleObjectMetadata>();
        var flatIndex = 0;

        for (var si = 0; si < presentation.Slides.Count; si++)
        {
            var slide = presentation.Slides[si];
            for (var shi = 0; shi < slide.Shapes.Count; shi++)
                if (slide.Shapes[shi] is IOleObjectFrame frame)
                {
                    items.Add(PptOleMetadataMapper.Map(frame, si, shi, flatIndex));
                    flatIndex++;
                }
        }

        return new OleListResult { Count = items.Count, Items = items };
    }
}
