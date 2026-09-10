namespace AsposeMcpServer.Helpers;

/// <summary>
///     The one set of limits on how large a table this server will build.
///     <para>
///         A table allocates an object per row-column pair, so the cost is the product and not
///         either dimension. Only the Word create path bounded it, with the numbers written inline;
///         the PDF and PowerPoint table handlers checked that each dimension was at least one and
///         nothing more, and splitting a Word cell had no bound at all — <c>splitCols</c> was
///         whatever the caller sent (R3-R04).
///     </para>
/// </summary>
public static class TableBudget
{
    /// <summary>Largest number of rows one table may have.</summary>
    public const int MaxRows = 10_000;

    /// <summary>
    ///     Largest number of columns one table may have. The Word format itself stops at 63, which
    ///     callers pass explicitly; the other formats have no such limit of their own.
    /// </summary>
    public const int MaxColumns = 1_000;

    /// <summary>Columns a Word table may have, which is a limit of the format.</summary>
    public const int MaxWordColumns = 63;

    /// <summary>Largest number of cells one table may hold.</summary>
    public const long MaxCells = 200_000;

    /// <summary>
    ///     Refuses a table whose dimensions, or whose number of cells, are above the budget.
    /// </summary>
    /// <param name="rows">Rows the table would have.</param>
    /// <param name="columns">Columns the table would have.</param>
    /// <param name="maxColumns">
    ///     Column limit to apply, so a format with a stricter one of its own can pass it.
    /// </param>
    /// <param name="what">What is being built, for the error message.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when either dimension is below one or above its limit, or when the product is
    ///     above <see cref="MaxCells" />.
    /// </exception>
    public static void EnsureWithinBudget(long rows, long columns, int maxColumns = MaxColumns,
        string what = "table")
    {
        if (rows < 1)
            throw new ArgumentException($"rows must be at least 1 for a {what}.", nameof(rows));
        if (columns < 1)
            throw new ArgumentException($"columns must be at least 1 for a {what}.", nameof(columns));

        if (rows > MaxRows)
            throw new ArgumentException(
                $"A {what} may have at most {MaxRows:N0} rows, but {rows:N0} were requested.",
                nameof(rows));
        if (columns > maxColumns)
            throw new ArgumentException(
                $"A {what} may have at most {maxColumns:N0} columns, but {columns:N0} were requested.",
                nameof(columns));

        // Checked in long: two values that each pass on their own multiply to something that does
        // not, and the product is what is actually allocated.
        var cells = rows * columns;
        if (cells > MaxCells)
            throw new ArgumentException(
                $"A {what} of {rows:N0} x {columns:N0} is {cells:N0} cells, above the limit of "
                + $"{MaxCells:N0}. Build a smaller table, or several.");
    }

    /// <summary>
    ///     Refuses a table whose rows are not all the same width.
    ///     <para>
    ///         A Word table may be ragged, so its cost is neither the product of two dimensions nor
    ///         readable from any one row. Pricing a split by the selected row's width let a table
    ///         whose <em>other</em> rows were near the column limit through, and undercounted the
    ///         cells of every table that was not rectangular (R4-R04).
    ///     </para>
    /// </summary>
    /// <param name="rows">How many rows the table will have.</param>
    /// <param name="maxColumns">The widest row the table will have, in cells.</param>
    /// <param name="totalCells">How many cells the table will hold in total.</param>
    /// <param name="columnLimit">The per-format column limit to apply.</param>
    /// <param name="what">What is being built, for the message.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the table would have no rows or columns, more than <see cref="MaxRows" />
    ///     rows, a row wider than <paramref name="columnLimit" />, or more than
    ///     <see cref="MaxCells" /> cells.
    /// </exception>
    public static void EnsureRaggedTableWithinBudget(long rows, long maxColumns, long totalCells,
        int columnLimit, string what)
    {
        if (rows < 1)
            throw new ArgumentException($"rows must be at least 1 for a {what}.", nameof(rows));
        if (maxColumns < 1)
            throw new ArgumentException($"columns must be at least 1 for a {what}.", nameof(maxColumns));

        if (rows > MaxRows)
            throw new ArgumentException(
                $"A {what} may have at most {MaxRows:N0} rows, but {rows:N0} were requested.",
                nameof(rows));
        if (maxColumns > columnLimit)
            throw new ArgumentException(
                $"A {what} may have at most {columnLimit:N0} columns, but {maxColumns:N0} were requested.",
                nameof(maxColumns));

        if (totalCells > MaxCells)
            throw new ArgumentException(
                $"A {what} of {rows:N0} rows would hold {totalCells:N0} cells, above the limit of "
                + $"{MaxCells:N0}. Build a smaller table, or several.");
    }
}
