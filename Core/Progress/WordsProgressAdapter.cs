using Aspose.Words.Saving;
using ModelContextProtocol;

namespace AsposeMcpServer.Core.Progress;

/// <summary>
///     Adapter for Aspose.Words progress reporting.
///     Implements <see cref="IDocumentSavingCallback" /> to convert Aspose progress
///     into MCP SDK progress notifications.
/// </summary>
public class WordsProgressAdapter : IDocumentSavingCallback
{
    /// <summary>
    ///     The MCP progress reporter.
    /// </summary>
    private readonly IProgress<ProgressNotificationValue>? _progress;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WordsProgressAdapter" /> class.
    /// </summary>
    /// <param name="progress">The MCP progress reporter to send notifications to.</param>
    public WordsProgressAdapter(IProgress<ProgressNotificationValue>? progress)
    {
        _progress = progress;
    }

    /// <summary>
    ///     Called by Aspose.Words during document saving operations.
    /// </summary>
    /// <param name="args">The document saving arguments containing progress information.</param>
    public void Notify(DocumentSavingArgs args)
    {
        // EstimatedProgress: 0.0 ~ 1.0
        var percentage = (int)(args.EstimatedProgress * 100);
        _progress?.Report(new ProgressNotificationValue { Progress = percentage, Total = 100 });
    }
}
