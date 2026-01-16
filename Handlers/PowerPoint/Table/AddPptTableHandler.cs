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
        var tableParams = ExtractTableParameters(parameters);
        ValidateTableParameters(tableParams);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, tableParams.SlideIndex);

        var table = CreateTable(slide, tableParams);
        PopulateTableCells(table, tableParams);

        MarkModified(context);

        var shapeIndex = slide.Shapes.Count - 1;
        return Success(
            $"Table added to slide {tableParams.SlideIndex} with {tableParams.Rows} rows and {tableParams.Columns} columns (shapeIndex: {shapeIndex}).");
    }

    /// <summary>
    ///     Extracts table parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted table parameters.</returns>
    private static TableParameters ExtractTableParameters(OperationParameters parameters)
    {
        return new TableParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("rows"),
            parameters.GetRequired<int>("columns"),
            parameters.GetOptional("x", 100.0),
            parameters.GetOptional("y", 100.0),
            parameters.GetOptional("columnWidth", 100.0),
            parameters.GetOptional("rowHeight", 30.0),
            parameters.GetOptional<string?>("data")
        );
    }

    /// <summary>
    ///     Validates the table parameters.
    /// </summary>
    /// <param name="p">The table parameters to validate.</param>
    private static void ValidateTableParameters(TableParameters p)
    {
        if (p.Rows < 1) throw new ArgumentException("rows must be at least 1");
        if (p.Columns < 1) throw new ArgumentException("columns must be at least 1");
    }

    /// <summary>
    ///     Creates a table on the slide.
    /// </summary>
    /// <param name="slide">The slide to add the table to.</param>
    /// <param name="p">The table parameters.</param>
    /// <returns>The created table.</returns>
    private static ITable CreateTable(ISlide slide, TableParameters p)
    {
        var colWidths = Enumerable.Repeat(p.ColumnWidth, p.Columns).ToArray();
        var rowHeights = Enumerable.Repeat(p.RowHeight, p.Rows).ToArray();
        return slide.Shapes.AddTable((float)p.X, (float)p.Y, colWidths, rowHeights);
    }

    /// <summary>
    ///     Populates the table cells with data.
    /// </summary>
    /// <param name="table">The table to populate.</param>
    /// <param name="p">The table parameters containing data.</param>
    private static void PopulateTableCells(ITable table, TableParameters p)
    {
        if (!string.IsNullOrEmpty(p.DataJson))
        {
            var data = JsonSerializer.Deserialize<string?[][]>(p.DataJson);
            if (data != null)
                for (var row = 0; row < Math.Min(p.Rows, data.Length); row++)
                for (var col = 0; col < Math.Min(p.Columns, data[row].Length); col++)
                    table[col, row].TextFrame.Text = data[row][col] ?? string.Empty;
        }
        else
        {
            for (var row = 0; row < p.Rows; row++)
            for (var col = 0; col < p.Columns; col++)
                table[col, row].TextFrame.Text = string.Empty;
        }
    }

    /// <summary>
    ///     Record for holding table creation parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="Rows">The number of rows.</param>
    /// <param name="Columns">The number of columns.</param>
    /// <param name="X">The X position.</param>
    /// <param name="Y">The Y position.</param>
    /// <param name="ColumnWidth">The column width.</param>
    /// <param name="RowHeight">The row height.</param>
    /// <param name="DataJson">The optional data JSON string.</param>
    private record TableParameters(
        int SlideIndex,
        int Rows,
        int Columns,
        double X,
        double Y,
        double ColumnWidth,
        double RowHeight,
        string? DataJson);
}
