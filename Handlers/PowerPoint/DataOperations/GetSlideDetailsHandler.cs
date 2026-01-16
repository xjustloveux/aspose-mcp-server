using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.DataOperations;

/// <summary>
///     Handler for getting detailed information about a slide.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSlideDetailsParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        var transition = slide.SlideShowTransition;
        object? transitionInfo = transition != null
            ? new
            {
                type = transition.Type.ToString(),
                speed = transition.Speed.ToString(),
                advanceOnClick = transition.AdvanceOnClick,
                advanceAfterTimeMs = transition.AdvanceAfterTime
            }
            : null;

        var animations = slide.Timeline.MainSequence;
        List<object> animationsList = [];
        for (var i = 0; i < animations.Count; i++)
        {
            var anim = animations[i];
            animationsList.Add(new
            {
                index = i,
                type = anim.Type.ToString(),
                targetShape = anim.TargetShape?.GetType().Name
            });
        }

        var background = slide.Background;
        object? backgroundInfo = background != null
            ? new { fillType = background.FillFormat.FillType.ToString() }
            : null;

        var notesSlide = slide.NotesSlideManager.NotesSlide;
        var notesText = notesSlide?.NotesTextFrame?.Text;

        string? thumbnailBase64 = null;
        if (p.IncludeThumbnail) thumbnailBase64 = PowerPointHelper.GenerateThumbnail(slide);

        var result = new
        {
            slideIndex = p.SlideIndex,
            hidden = slide.Hidden,
            layout = slide.LayoutSlide?.Name,
            slideSize = new
            {
                width = presentation.SlideSize.Size.Width,
                height = presentation.SlideSize.Size.Height
            },
            shapesCount = slide.Shapes.Count,
            transition = transitionInfo,
            animationsCount = animations.Count,
            animations = animationsList,
            background = backgroundInfo,
            notes = string.IsNullOrWhiteSpace(notesText) ? null : notesText,
            thumbnail = thumbnailBase64
        };

        return JsonResult(result);
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
    private record SlideDetailsParameters(int SlideIndex, bool IncludeThumbnail);
}
