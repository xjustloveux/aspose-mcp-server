using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Layout;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for getting available layouts from a PowerPoint presentation.
/// </summary>
[ResultType(typeof(GetLayoutsResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetLayoutsParameters(parameters);

        var presentation = context.Document;

        if (p.MasterIndex.HasValue)
        {
            PowerPointHelper.ValidateCollectionIndex(p.MasterIndex.Value, presentation.Masters.Count, "master");

            var master = presentation.Masters[p.MasterIndex.Value];
            var layoutsList = PptLayoutHelper.BuildLayoutsList(master.LayoutSlides);

            var result = new GetLayoutsResult
            {
                MasterIndex = p.MasterIndex.Value,
                Count = master.LayoutSlides.Count,
                Layouts = layoutsList
            };

            return result;
        }
        else
        {
            List<GetLayoutMasterInfo> mastersList = [];

            for (var i = 0; i < presentation.Masters.Count; i++)
            {
                var master = presentation.Masters[i];
                mastersList.Add(new GetLayoutMasterInfo
                {
                    MasterIndex = i,
                    LayoutCount = master.LayoutSlides.Count,
                    Layouts = PptLayoutHelper.BuildLayoutsList(master.LayoutSlides)
                });
            }

            var result = new GetLayoutsResult
            {
                MastersCount = presentation.Masters.Count,
                Masters = mastersList
            };

            return result;
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
    private sealed record GetLayoutsParameters(int? MasterIndex);
}
