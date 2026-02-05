using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Watermark;

/// <summary>
///     Handler for removing watermarks from a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemovePptWatermarkHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes all watermarks from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>Success message with removal count.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        var removedCount = 0;

        foreach (var slide in presentation.Slides)
        {
            var shapesToRemove = new List<IShape>();
            foreach (var shape in slide.Shapes)
                if (shape.Name != null &&
                    shape.Name.StartsWith(AddTextPptWatermarkHandler.WatermarkPrefix, StringComparison.Ordinal))
                    shapesToRemove.Add(shape);

            foreach (var shape in shapesToRemove)
            {
                slide.Shapes.Remove(shape);
                removedCount++;
            }
        }

        if (removedCount > 0)
            MarkModified(context);

        return new SuccessResult { Message = $"Removed {removedCount} watermark(s) from presentation." };
    }
}
