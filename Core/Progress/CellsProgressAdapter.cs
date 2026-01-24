using Aspose.Cells.Rendering;
using ModelContextProtocol;

namespace AsposeMcpServer.Core.Progress;

/// <summary>
///     Adapter for Aspose.Cells progress reporting.
///     Implements <see cref="IPageSavingCallback" /> to convert Aspose progress
///     into MCP SDK progress notifications.
/// </summary>
public class CellsProgressAdapter : IPageSavingCallback
{
    /// <summary>
    ///     The MCP progress reporter.
    /// </summary>
    private readonly IProgress<ProgressNotificationValue>? _progress;

    /// <summary>
    ///     Initializes a new instance of the <see cref="CellsProgressAdapter" /> class.
    /// </summary>
    /// <param name="progress">The MCP progress reporter to send notifications to.</param>
    public CellsProgressAdapter(IProgress<ProgressNotificationValue>? progress)
    {
        _progress = progress;
    }

    /// <summary>
    ///     Called when a page starts saving.
    /// </summary>
    /// <param name="args">The page start saving arguments.</param>
    public void PageStartSaving(PageStartSavingArgs args)
    {
        if (args.PageCount > 0)
        {
            var percentage = args.PageIndex * 100 / args.PageCount;
            _progress?.Report(new ProgressNotificationValue { Progress = percentage, Total = 100 });
        }
    }

    /// <summary>
    ///     Called when a page finishes saving.
    /// </summary>
    /// <param name="args">The page end saving arguments.</param>
    public void PageEndSaving(PageEndSavingArgs args)
    {
        if (args.PageCount > 0)
        {
            var percentage = (args.PageIndex + 1) * 100 / args.PageCount;
            _progress?.Report(new ProgressNotificationValue { Progress = percentage, Total = 100 });
        }
    }
}
