using Aspose.Slides;
using ModelContextProtocol;

namespace AsposeMcpServer.Core.Progress;

/// <summary>
///     Adapter for Aspose.Slides progress reporting.
///     Implements <see cref="IProgressCallback" /> to convert Aspose progress
///     into MCP SDK progress notifications.
/// </summary>
public class SlidesProgressAdapter : IProgressCallback
{
    /// <summary>
    ///     The MCP progress reporter.
    /// </summary>
    private readonly IProgress<ProgressNotificationValue>? _progress;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SlidesProgressAdapter" /> class.
    /// </summary>
    /// <param name="progress">The MCP progress reporter to send notifications to.</param>
    public SlidesProgressAdapter(IProgress<ProgressNotificationValue>? progress)
    {
        _progress = progress;
    }

    /// <summary>
    ///     Called by Aspose.Slides during presentation operations.
    /// </summary>
    /// <param name="progressValue">The progress value from 0 to 100.</param>
    public void Reporting(double progressValue)
    {
        // progressValue: 0 ~ 100
        _progress?.Report(new ProgressNotificationValue { Progress = (int)progressValue, Total = 100 });
    }
}
