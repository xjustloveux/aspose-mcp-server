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
        var editParams = ExtractEditTableParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, editParams.SlideIndex);
        var table = PptTableHelper.GetTable(slide, editParams.ShapeIndex);

        if (editParams.X.HasValue)
            table.X = editParams.X.Value;

        if (editParams.Y.HasValue)
            table.Y = editParams.Y.Value;

        if (editParams.Width.HasValue)
            table.Width = editParams.Width.Value;

        if (editParams.Height.HasValue)
            table.Height = editParams.Height.Value;

        MarkModified(context);

        return Success($"Table properties updated on slide {editParams.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts edit table parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit table parameters.</returns>
    private static EditTableParameters ExtractEditTableParameters(OperationParameters parameters)
    {
        return new EditTableParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<float?>("x"),
            parameters.GetOptional<float?>("y"),
            parameters.GetOptional<float?>("width"),
            parameters.GetOptional<float?>("height")
        );
    }

    /// <summary>
    ///     Record for holding edit table parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="X">The optional X position.</param>
    /// <param name="Y">The optional Y position.</param>
    /// <param name="Width">The optional width.</param>
    /// <param name="Height">The optional height.</param>
    private sealed record EditTableParameters(
        int SlideIndex,
        int ShapeIndex,
        float? X,
        float? Y,
        float? Width,
        float? Height);
}
