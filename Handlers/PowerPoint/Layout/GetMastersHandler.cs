using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for getting master slides information from a PowerPoint presentation.
/// </summary>
public class GetMastersHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_masters";

    /// <summary>
    ///     Gets master slides information.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing master slide information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;

        if (presentation.Masters.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                masters = Array.Empty<object>(),
                message = "No master slides found"
            };
            return JsonResult(emptyResult);
        }

        List<object> mastersList = [];

        for (var i = 0; i < presentation.Masters.Count; i++)
        {
            var master = presentation.Masters[i];
            mastersList.Add(new
            {
                index = i,
                name = master.Name,
                layoutCount = master.LayoutSlides.Count,
                layouts = PptLayoutHelper.BuildLayoutsList(master.LayoutSlides)
            });
        }

        var result = new
        {
            count = presentation.Masters.Count,
            masters = mastersList
        };

        return JsonResult(result);
    }
}
