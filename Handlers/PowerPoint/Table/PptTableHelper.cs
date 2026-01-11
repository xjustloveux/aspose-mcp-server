using Aspose.Slides;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Helper class providing shared methods for PowerPoint table handlers.
/// </summary>
public static class PptTableHelper
{
    /// <summary>
    ///     Gets a slide from the presentation by index with validation.
    /// </summary>
    /// <param name="presentation">The presentation to get the slide from.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <returns>The slide at the specified index.</returns>
    /// <exception cref="ArgumentException">Thrown when the slide index is out of range.</exception>
    public static ISlide GetSlide(Presentation presentation, int slideIndex)
    {
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException(
                $"slideIndex must be between 0 and {presentation.Slides.Count - 1}, got: {slideIndex}");

        return presentation.Slides[slideIndex];
    }

    /// <summary>
    ///     Gets a table from the slide by shape index with validation.
    /// </summary>
    /// <param name="slide">The slide to get the table from.</param>
    /// <param name="shapeIndex">The zero-based index of the shape.</param>
    /// <returns>The table at the specified shape index.</returns>
    /// <exception cref="ArgumentException">Thrown when the shape index is out of range or the shape is not a table.</exception>
    public static ITable GetTable(ISlide slide, int shapeIndex)
    {
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
            throw new ArgumentException(
                $"shapeIndex must be between 0 and {slide.Shapes.Count - 1}, got: {shapeIndex}");

        if (slide.Shapes[shapeIndex] is not ITable table)
            throw new ArgumentException($"Shape at index {shapeIndex} is not a table");

        return table;
    }

    /// <summary>
    ///     Validates that the row index is within the valid range for the table.
    /// </summary>
    /// <param name="table">The table to validate against.</param>
    /// <param name="rowIndex">The row index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the row index is out of range.</exception>
    public static void ValidateRowIndex(ITable table, int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
            throw new ArgumentException(
                $"rowIndex must be between 0 and {table.Rows.Count - 1}, got: {rowIndex}");
    }

    /// <summary>
    ///     Validates that the column index is within the valid range for the table.
    /// </summary>
    /// <param name="table">The table to validate against.</param>
    /// <param name="columnIndex">The column index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the column index is out of range.</exception>
    public static void ValidateColumnIndex(ITable table, int columnIndex)
    {
        if (columnIndex < 0 || columnIndex >= table.Columns.Count)
            throw new ArgumentException(
                $"columnIndex must be between 0 and {table.Columns.Count - 1}, got: {columnIndex}");
    }
}
