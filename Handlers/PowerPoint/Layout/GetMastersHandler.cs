using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Layout;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for getting master slides information from a PowerPoint presentation.
/// </summary>
[ResultType(typeof(GetMastersResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;

        if (presentation.Masters.Count == 0)
        {
            var emptyResult = new GetMastersResult
            {
                Count = 0,
                Masters = [],
                Message = "No master slides found"
            };
            return emptyResult;
        }

        List<GetMasterInfo> mastersList = [];

        for (var i = 0; i < presentation.Masters.Count; i++)
        {
            var master = presentation.Masters[i];
            mastersList.Add(new GetMasterInfo
            {
                Index = i,
                Name = master.Name,
                LayoutCount = master.LayoutSlides.Count,
                Layouts = PptLayoutHelper.BuildLayoutsList(master.LayoutSlides)
            });
        }

        var result = new GetMastersResult
        {
            Count = presentation.Masters.Count,
            Masters = mastersList
        };

        return result;
    }
}
