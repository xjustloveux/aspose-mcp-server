using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     A table costs one object per row-column pair, so the product is what has to be bounded.
///     Only the Word create path did that; the PDF and PowerPoint handlers checked that each
///     dimension was at least one and nothing more, and splitting a Word cell had no bound at all
///     (R3-R04).
/// </summary>
public class TableBudgetTests
{
    /// <summary>
    ///     R4-R04: a ragged table's cost is a sum, not a product. Pricing one row's width as the
    ///     whole table's undercounted every table that is not rectangular.
    /// </summary>
    [Fact]
    public void EnsureRaggedTableWithinBudget_ShouldCountEveryRow()
    {
        // 3,340 rows of 60 cells and one of 1: the product of the narrow row's width and the row
        // count is 6,682, well inside the cap, while the table actually holds 200,401 cells.
        const long rows = 3_341;
        const long widest = 60;
        const long overTheCap = TableBudget.MaxCells + 401;

        var refusal = Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureRaggedTableWithinBudget(rows, widest, overTheCap,
                TableBudget.MaxWordColumns, "split cell"));
        Assert.Contains(TableBudget.MaxCells.ToString("N0"), refusal.Message, StringComparison.Ordinal);

        TableBudget.EnsureRaggedTableWithinBudget(rows, widest, TableBudget.MaxCells,
            TableBudget.MaxWordColumns, "split cell");
    }

    [Fact]
    public void EnsureRaggedTableWithinBudget_ShouldRefuseTheWidestRowNotTheAverage()
    {
        Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureRaggedTableWithinBudget(2, TableBudget.MaxWordColumns + 1, 100,
                TableBudget.MaxWordColumns, "split cell"));

        TableBudget.EnsureRaggedTableWithinBudget(2, TableBudget.MaxWordColumns, 100,
            TableBudget.MaxWordColumns, "split cell");
    }

    [Theory]
    [InlineData(0, 5, 5)]
    [InlineData(5, 0, 5)]
    [InlineData(TableBudget.MaxRows + 1, 5, 5)]
    public void EnsureRaggedTableWithinBudget_ShouldRefuseAnImpossibleShape(long rows, long widest,
        long cells)
    {
        Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureRaggedTableWithinBudget(rows, widest, cells,
                TableBudget.MaxWordColumns, "split cell"));
    }

    [Fact]
    public void EnsureWithinBudget_WithAnOrdinaryTable_ShouldBeAccepted()
    {
        var exception = Record.Exception(() => TableBudget.EnsureWithinBudget(20, 8));

        Assert.Null(exception);
    }

    /// <param name="rows">Rows requested.</param>
    /// <param name="columns">Columns requested.</param>
    [Theory]
    [InlineData(0, 4)]
    [InlineData(4, 0)]
    [InlineData(-1, 4)]
    [InlineData(4, -1)]
    public void EnsureWithinBudget_WithADimensionBelowOne_ShouldBeRefused(long rows, long columns)
    {
        Assert.Throws<ArgumentException>(() => TableBudget.EnsureWithinBudget(rows, columns));
    }

    [Fact]
    public void EnsureWithinBudget_AtTheRowLimit_ShouldBeAccepted()
    {
        var exception = Record.Exception(() =>
            TableBudget.EnsureWithinBudget(TableBudget.MaxRows, 1));

        Assert.Null(exception);
    }

    [Fact]
    public void EnsureWithinBudget_OneRowPastTheLimit_ShouldBeRefused()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureWithinBudget(TableBudget.MaxRows + 1, 1));

        Assert.Contains("rows", exception.Message);
    }

    [Fact]
    public void EnsureWithinBudget_AtTheColumnLimit_ShouldBeAccepted()
    {
        var exception = Record.Exception(() =>
            TableBudget.EnsureWithinBudget(1, TableBudget.MaxColumns));

        Assert.Null(exception);
    }

    [Fact]
    public void EnsureWithinBudget_OneColumnPastTheLimit_ShouldBeRefused()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureWithinBudget(1, TableBudget.MaxColumns + 1));

        Assert.Contains("columns", exception.Message);
    }

    /// <summary>
    ///     The case neither per-axis check can see: both dimensions are legal on their own.
    /// </summary>
    [Fact]
    public void EnsureWithinBudget_WithTooManyCells_ShouldBeRefused()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureWithinBudget(5_000, 1_000));

        Assert.Contains("cells", exception.Message);
    }

    [Fact]
    public void EnsureWithinBudget_AtExactlyTheCellLimit_ShouldBeAccepted()
    {
        var exception = Record.Exception(() =>
            TableBudget.EnsureWithinBudget(TableBudget.MaxCells / TableBudget.MaxColumns,
                TableBudget.MaxColumns));

        Assert.Null(exception);
    }

    [Fact]
    public void EnsureWithinBudget_OneCellPastTheLimit_ShouldBeRefused()
    {
        Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureWithinBudget(TableBudget.MaxCells / TableBudget.MaxColumns + 1,
                TableBudget.MaxColumns));
    }

    /// <summary>
    ///     Word tables stop at 63 columns because the format does, and callers pass that limit.
    /// </summary>
    [Fact]
    public void EnsureWithinBudget_WithTheWordColumnLimit_ShouldApplyIt()
    {
        TableBudget.EnsureWithinBudget(10, TableBudget.MaxWordColumns, TableBudget.MaxWordColumns);

        Assert.Throws<ArgumentException>(() =>
            TableBudget.EnsureWithinBudget(10, TableBudget.MaxWordColumns + 1,
                TableBudget.MaxWordColumns));
    }
}
