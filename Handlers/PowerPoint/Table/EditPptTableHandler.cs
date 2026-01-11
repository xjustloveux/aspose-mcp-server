using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for editing table properties in PowerPoint presentations.
/// </summary>
public class EditPptTableHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits table properties such as position and size.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0), x, y, width, height.
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var x = parameters.GetOptional<float?>("x");
        var y = parameters.GetOptional<float?>("y");
        var width = parameters.GetOptional<float?>("width");
        var height = parameters.GetOptional<float?>("height");

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, slideIndex);
        var table = PptTableHelper.GetTable(slide, shapeIndex);

        if (x.HasValue)
            table.X = x.Value;

        if (y.HasValue)
            table.Y = y.Value;

        if (width.HasValue)
            table.Width = width.Value;

        if (height.HasValue)
            table.Height = height.Value;

        MarkModified(context);

        return Success($"Table properties updated on slide {slideIndex}.");
    }
}
