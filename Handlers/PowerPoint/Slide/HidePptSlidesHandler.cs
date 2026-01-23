using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for hiding or showing slides in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class HidePptSlidesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "hide";

    /// <summary>
    ///     Hides or shows slides in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: hidden (true to hide, false to show)
    ///     Optional: slideIndices (JSON array of indices, default: all slides)
    /// </param>
    /// <returns>Success message with hide/show details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractHidePptSlidesParameters(parameters);

        var presentation = context.Document;

        int[] targets;
        if (!string.IsNullOrWhiteSpace(p.SlideIndicesJson))
        {
            var indices = JsonSerializer.Deserialize<int[]>(p.SlideIndicesJson);
            targets = indices ?? Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }
        else
        {
            targets = Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        foreach (var idx in targets)
            presentation.Slides[idx].Hidden = p.Hidden;

        MarkModified(context);

        return new SuccessResult { Message = $"Set {targets.Length} slide(s) hidden={p.Hidden}." };
    }

    /// <summary>
    ///     Extracts parameters for hide slides operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static HidePptSlidesParameters ExtractHidePptSlidesParameters(OperationParameters parameters)
    {
        return new HidePptSlidesParameters(
            parameters.GetOptional("hidden", false),
            parameters.GetOptional<string?>("slideIndices"));
    }

    /// <summary>
    ///     Parameters for hide slides operation.
    /// </summary>
    /// <param name="Hidden">Whether to hide or show the slides.</param>
    /// <param name="SlideIndicesJson">JSON array of slide indices.</param>
    private sealed record HidePptSlidesParameters(bool Hidden, string? SlideIndicesJson);
}
