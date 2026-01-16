using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for getting table content from PowerPoint presentations.
/// </summary>
public class GetPptTableContentHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_content";

    /// <summary>
    ///     Gets the content of a table.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>JSON result with table content.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var getParams = ExtractGetContentParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, getParams.SlideIndex);
        var table = PptTableHelper.GetTable(slide, getParams.ShapeIndex);

        var rows = new List<List<string>>();
        for (var row = 0; row < table.Rows.Count; row++)
        {
            var rowData = new List<string>();
            for (var col = 0; col < table.Columns.Count; col++)
            {
                var cellText = table[col, row].TextFrame?.Text ?? string.Empty;
                rowData.Add(cellText);
            }

            rows.Add(rowData);
        }

        var result = new
        {
            slideIndex = getParams.SlideIndex,
            shapeIndex = getParams.ShapeIndex,
            rowCount = table.Rows.Count,
            columnCount = table.Columns.Count,
            data = rows
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts get content parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get content parameters.</returns>
    private static GetContentParameters ExtractGetContentParameters(OperationParameters parameters)
    {
        return new GetContentParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex")
        );
    }

    /// <summary>
    ///     Record for holding get content parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    private sealed record GetContentParameters(int SlideIndex, int ShapeIndex);
}
