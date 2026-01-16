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
        var p = ExtractGetLayoutsParameters(parameters);

        var presentation = context.Document;

        if (p.MasterIndex.HasValue)
        {
            PowerPointHelper.ValidateCollectionIndex(p.MasterIndex.Value, presentation.Masters.Count, "master");

            var master = presentation.Masters[p.MasterIndex.Value];
            var layoutsList = PptLayoutHelper.BuildLayoutsList(master.LayoutSlides);

            var result = new
            {
                masterIndex = p.MasterIndex.Value,
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

    /// <summary>
    ///     Extracts get layouts parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get layouts parameters.</returns>
    private static GetLayoutsParameters ExtractGetLayoutsParameters(OperationParameters parameters)
    {
        return new GetLayoutsParameters(parameters.GetOptional<int?>("masterIndex"));
    }

    /// <summary>
    ///     Record for holding get layouts parameters.
    /// </summary>
    /// <param name="MasterIndex">The optional master index.</param>
    private record GetLayoutsParameters(int? MasterIndex);
}
