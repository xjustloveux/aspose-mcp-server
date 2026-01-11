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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var includeThumbnail = parameters.GetOptional<bool>("includeThumbnail");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

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
        if (includeThumbnail) thumbnailBase64 = PowerPointHelper.GenerateThumbnail(slide);

        var result = new
        {
            slideIndex,
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
}
