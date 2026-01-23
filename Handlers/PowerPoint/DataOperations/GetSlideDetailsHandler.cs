using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.DataOperations;

namespace AsposeMcpServer.Handlers.PowerPoint.DataOperations;

/// <summary>
///     Handler for getting detailed information about a slide.
/// </summary>
[ResultType(typeof(GetSlideDetailsResult))]
public class GetSlideDetailsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_slide_details";

    /// <summary>
    ///     Gets detailed information about a specific slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex
    ///     Optional: includeThumbnail (default false)
    /// </param>
    /// <returns>JSON string containing detailed slide information including layout, transitions, and animations.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSlideDetailsParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        var transition = slide.SlideShowTransition;
        var transitionInfo = transition != null
            ? new GetSlideDetailsTransitionInfo
            {
                Type = transition.Type.ToString(),
                Speed = transition.Speed.ToString(),
                AdvanceOnClick = transition.AdvanceOnClick,
                AdvanceAfterTimeMs = transition.AdvanceAfterTime
            }
            : null;

        var animations = slide.Timeline.MainSequence;
        List<GetSlideDetailsAnimationInfo> animationsList = [];
        for (var i = 0; i < animations.Count; i++)
        {
            var anim = animations[i];
            animationsList.Add(new GetSlideDetailsAnimationInfo
            {
                Index = i,
                Type = anim.Type.ToString(),
                TargetShape = anim.TargetShape?.GetType().Name
            });
        }

        var background = slide.Background;
        var backgroundInfo = background != null
            ? new GetSlideDetailsBackgroundInfo { FillType = background.FillFormat.FillType.ToString() }
            : null;

        var notesSlide = slide.NotesSlideManager.NotesSlide;
        var notesText = notesSlide?.NotesTextFrame?.Text;

        string? thumbnailBase64 = null;
        if (p.IncludeThumbnail) thumbnailBase64 = PowerPointHelper.GenerateThumbnail(slide);

        var result = new GetSlideDetailsResult
        {
            SlideIndex = p.SlideIndex,
            Hidden = slide.Hidden,
            Layout = slide.LayoutSlide?.Name,
            SlideSize = new GetSlideDetailsSizeInfo
            {
                Width = presentation.SlideSize.Size.Width,
                Height = presentation.SlideSize.Size.Height
            },
            ShapesCount = slide.Shapes.Count,
            Transition = transitionInfo,
            AnimationsCount = animations.Count,
            Animations = animationsList,
            Background = backgroundInfo,
            Notes = string.IsNullOrWhiteSpace(notesText) ? null : notesText,
            Thumbnail = thumbnailBase64
        };

        return result;
    }

    /// <summary>
    ///     Extracts slide details parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted slide details parameters.</returns>
    private static SlideDetailsParameters ExtractSlideDetailsParameters(OperationParameters parameters)
    {
        return new SlideDetailsParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetOptional<bool>("includeThumbnail"));
    }

    /// <summary>
    ///     Record for holding slide details parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="IncludeThumbnail">Whether to include slide thumbnail.</param>
    private sealed record SlideDetailsParameters(int SlideIndex, bool IncludeThumbnail);
}
