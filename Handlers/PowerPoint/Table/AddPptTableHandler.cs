using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for adding tables to PowerPoint presentations.
/// </summary>
public class AddPptTableHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new table to a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: rows, columns.
    ///     Optional: slideIndex (default: 0), x, y, columnWidth, rowHeight, data.
    /// </param>
    /// <returns>Success message with table creation details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var rows = parameters.GetRequired<int>("rows");
        var columns = parameters.GetRequired<int>("columns");
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 100.0);
        var columnWidth = parameters.GetOptional("columnWidth", 100.0);
        var rowHeight = parameters.GetOptional("rowHeight", 30.0);
        var dataJson = parameters.GetOptional<string?>("data");

        if (rows < 1)
            throw new ArgumentException("rows must be at least 1");
        if (columns < 1)
            throw new ArgumentException("columns must be at least 1");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var colWidths = Enumerable.Repeat(columnWidth, columns).ToArray();
        var rowHeights = Enumerable.Repeat(rowHeight, rows).ToArray();

        var table = slide.Shapes.AddTable((float)x, (float)y, colWidths, rowHeights);

        if (!string.IsNullOrEmpty(dataJson))
        {
            var data = JsonSerializer.Deserialize<string?[][]>(dataJson);
            if (data != null)
                for (var row = 0; row < Math.Min(rows, data.Length); row++)
                for (var col = 0; col < Math.Min(columns, data[row].Length); col++)
                    table[col, row].TextFrame.Text = data[row][col] ?? string.Empty;
        }
        else
        {
            for (var row = 0; row < rows; row++)
            for (var col = 0; col < columns; col++)
                table[col, row].TextFrame.Text = string.Empty;
        }

        MarkModified(context);

        var shapeIndex = slide.Shapes.Count - 1;
        return Success(
            $"Table added to slide {slideIndex} with {rows} rows and {columns} columns (shapeIndex: {shapeIndex}).");
    }
}
