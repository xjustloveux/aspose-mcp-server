using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.PowerPoint.Watermark;

namespace AsposeMcpServer.Handlers.PowerPoint.Watermark;

/// <summary>
///     Handler for getting watermarks from a PowerPoint presentation.
/// </summary>
[ResultType(typeof(GetWatermarksPptResult))]
public class GetPptWatermarksHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all watermarks from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>List of watermarks found in the presentation.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        var watermarks = new List<PptWatermarkInfo>();

        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            foreach (var shape in slide.Shapes)
                if (shape.Name != null &&
                    shape.Name.StartsWith(AddTextPptWatermarkHandler.WatermarkPrefix, StringComparison.Ordinal))
                {
                    var isText = shape.Name.Contains("_TEXT_");
                    string? wmText = null;
                    if (isText && shape is IAutoShape { TextFrame: not null } autoShape)
                        wmText = autoShape.TextFrame.Text;

                    watermarks.Add(new PptWatermarkInfo
                    {
                        SlideIndex = i,
                        ShapeName = shape.Name,
                        Type = isText ? "text" : "image",
                        Text = wmText
                    });
                }
        }

        return new GetWatermarksPptResult
        {
            Count = watermarks.Count,
            Items = watermarks,
            Message = $"Found {watermarks.Count} watermark(s) in presentation."
        };
    }
}
