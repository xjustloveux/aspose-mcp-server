using ModelContextProtocol;

namespace AsposeMcpServer.Core.Progress;

/// <summary>
///     Adapter for Aspose.PDF progress reporting.
///     Provides manual progress reporting methods for PDF operations.
/// </summary>
/// <remarks>
///     Aspose.PDF 23.10 does not provide a built-in progress callback for document saving.
///     This adapter provides helper methods for reporting progress at key milestones
///     during PDF operations (merge, split, compress, etc.).
/// </remarks>
public class PdfProgressAdapter
{
    /// <summary>
    ///     The MCP progress reporter.
    /// </summary>
    private readonly IProgress<ProgressNotificationValue>? _progress;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfProgressAdapter" /> class.
    /// </summary>
    /// <param name="progress">The MCP progress reporter to send notifications to.</param>
    public PdfProgressAdapter(IProgress<ProgressNotificationValue>? progress)
    {
        _progress = progress;
    }

    /// <summary>
    ///     Reports progress for operations processing multiple items.
    /// </summary>
    /// <param name="currentItem">The current item being processed (0-based).</param>
    /// <param name="totalItems">The total number of items to process.</param>
    /// <param name="message">Optional message describing the current operation.</param>
    public void ReportProgress(int currentItem, int totalItems, string? message = null)
    {
        if (totalItems > 0)
        {
            var percentage = (currentItem + 1) * 100 / totalItems;
            _progress?.Report(new ProgressNotificationValue { Progress = percentage, Total = 100, Message = message });
        }
    }

    /// <summary>
    ///     Reports a specific percentage of progress.
    /// </summary>
    /// <param name="percentage">The progress percentage (0-100).</param>
    /// <param name="message">Optional message describing the current operation.</param>
    public void ReportPercentage(int percentage, string? message = null)
    {
        _progress?.Report(new ProgressNotificationValue { Progress = percentage, Total = 100, Message = message });
    }
}
