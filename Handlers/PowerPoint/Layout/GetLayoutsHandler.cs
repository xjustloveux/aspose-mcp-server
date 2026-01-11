using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for getting available layouts from a PowerPoint presentation.
/// </summary>
public class GetLayoutsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_layouts";

    /// <summary>
    ///     Gets available layouts with layout type information.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: masterIndex
    /// </param>
    /// <returns>JSON string containing layout information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var masterIndex = parameters.GetOptional<int?>("masterIndex");

        var presentation = context.Document;

        if (masterIndex.HasValue)
        {
            PowerPointHelper.ValidateCollectionIndex(masterIndex.Value, presentation.Masters.Count, "master");

            var master = presentation.Masters[masterIndex.Value];
            var layoutsList = PptLayoutHelper.BuildLayoutsList(master.LayoutSlides);

            var result = new
            {
                masterIndex = masterIndex.Value,
                count = master.LayoutSlides.Count,
                layouts = layoutsList
            };

            return JsonResult(result);
        }
        else
        {
            List<object> mastersList = [];

            for (var i = 0; i < presentation.Masters.Count; i++)
            {
                var master = presentation.Masters[i];
                mastersList.Add(new
                {
                    masterIndex = i,
                    layoutCount = master.LayoutSlides.Count,
                    layouts = PptLayoutHelper.BuildLayoutsList(master.LayoutSlides)
                });
            }

            var result = new
            {
                mastersCount = presentation.Masters.Count,
                masters = mastersList
            };

            return JsonResult(result);
        }
    }
}
