using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     A render was priced from an average page size, and a request for a single page was not
///     priced at all — so one very large page at a high resolution was an unbounded allocation
///     (R3-R03). These fixtures drive the running count that replaced the estimate.
/// </summary>
public class PixelBudgetTests
{
    [Fact]
    public void Add_WithAnOrdinaryPage_ShouldBeAccepted()
    {
        var budget = new PixelBudget();

        budget.Add(8.27, 11.69, 300);

        Assert.True(budget.Total > 0);
        Assert.True(budget.Total < RenderBudget.MaxTotalPixels);
    }

    /// <summary>
    ///     The case the single-page branches never checked: one sheet far larger than A4.
    /// </summary>
    [Fact]
    public void Add_WithOneEnormousPage_ShouldBeRefused()
    {
        var budget = new PixelBudget();

        var exception = Assert.Throws<ArgumentException>(() => budget.Add(200, 200, 3000));

        Assert.Contains("above the limit", exception.Message);
    }

    [Fact]
    public void Add_AtExactlyTheLimit_ShouldBeAccepted()
    {
        // 50,000 x 40,000 at 1 DPI is 2,000,000,000 pixels exactly, and both factors are
        // representable, so the fixture sits on the boundary rather than near it.
        var budget = new PixelBudget();

        budget.Add(50_000, 40_000, 1);

        Assert.Equal(RenderBudget.MaxTotalPixels, budget.Total);
    }

    [Fact]
    public void Add_OnePixelPastTheLimit_ShouldBeRefused()
    {
        var budget = new PixelBudget();

        Assert.Throws<ArgumentException>(() => budget.Add(50_000, 40_001, 1));
    }

    /// <summary>
    ///     Pages that are each acceptable can still be unacceptable together, which is the whole
    ///     reason the count runs rather than being taken one page at a time.
    /// </summary>
    [Fact]
    public void Add_AcrossManyPages_ShouldAccumulate()
    {
        var budget = new PixelBudget();

        var exception = Assert.Throws<ArgumentException>(() =>
        {
            for (var page = 0; page < 1000; page++)
                budget.Add(8.27, 11.69, 600);
        });

        Assert.Contains("page(s)", exception.Message);
    }

    /// <summary>
    ///     A product this large overflows a 64-bit integer, and an overflowed total compares as
    ///     smaller than the limit rather than larger.
    /// </summary>
    [Fact]
    public void Add_WithAValueThatWouldOverflowALong_ShouldStillBeRefused()
    {
        var budget = new PixelBudget();

        Assert.Throws<ArgumentException>(() => budget.Add(1_000_000, 1_000_000, 1200));
    }

    /// <param name="width">Width in inches; zero or negative means unknown.</param>
    /// <param name="height">Height in inches; zero or negative means unknown.</param>
    [Theory]
    [InlineData(0, 0)]
    [InlineData(-1, -1)]
    public void Add_WithAnUnknownPageSize_ShouldFallBackToA4(double width, double height)
    {
        var budget = new PixelBudget();
        var reference = new PixelBudget();

        budget.Add(width, height, 300);
        reference.Add(RenderBudget.DefaultPageWidthInches, RenderBudget.DefaultPageHeightInches, 300);

        Assert.Equal(reference.Total, budget.Total);
    }
}
