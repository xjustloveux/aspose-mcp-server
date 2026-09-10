using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Covers A-06's aggregate half. Every individual parameter was already capped — DPI at 1200,
///     input arrays at 1000, page counts per operation — but the cost of a request is their
///     product. A 500-page document at 1200 DPI breaks no single limit and still asks for roughly
///     a hundred billion pixels, so the bound has to be on the total.
/// </summary>
public class RenderBudgetTests
{
    [Fact]
    public void AnOrdinaryConversion_ShouldBeAllowed()
    {
        // 50 A4 pages at 150 DPI: about 87 million pixels.
        var exception = Record.Exception(() => RenderBudget.EnsureWithinBudget(50, 150));

        Assert.Null(exception);
    }

    [Fact]
    public void ALargeButReasonableConversion_ShouldBeAllowed()
    {
        // 200 pages at 300 DPI: about 1.4 billion pixels, still inside the budget.
        var exception = Record.Exception(() => RenderBudget.EnsureWithinBudget(200, 300));

        Assert.Null(exception);
    }

    [Fact]
    public void ManyPagesAtHighDpi_ShouldBeRefused()
    {
        // Each parameter is individually legal; the product is not.
        var exception = Assert.Throws<ArgumentException>(() => RenderBudget.EnsureWithinBudget(500, 1200));

        Assert.Contains("pixels", exception.Message);
        Assert.Contains("1200 DPI", exception.Message);
    }

    [Fact]
    public void TooManyOutputFiles_ShouldBeRefusedEvenAtLowDpi()
    {
        var exception = Assert.Throws<ArgumentException>(() => RenderBudget.EnsureWithinBudget(10_000, 10));

        Assert.Contains("image files", exception.Message);
    }

    [Fact]
    public void TheFileLimitShouldBeReportedBeforeThePixelLimit()
    {
        // Both are exceeded here; naming the file count first is the more actionable message.
        var exception = Assert.Throws<ArgumentException>(() => RenderBudget.EnsureWithinBudget(20_000, 600));

        Assert.Contains("image files", exception.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void ANonPositivePageCount_ShouldNotBeJudged(int pageCount)
    {
        // Nothing will be rendered, so there is nothing to budget; the caller's own validation
        // reports an empty document.
        var exception = Record.Exception(() => RenderBudget.EnsureWithinBudget(pageCount, 1200));

        Assert.Null(exception);
    }

    [Fact]
    public void ASinglePageAtTheHighestAllowedDpi_ShouldBeAllowed()
    {
        // The per-parameter DPI cap is 1200; one page at that resolution must remain possible,
        // otherwise the aggregate budget would silently contradict it.
        var exception = Record.Exception(() => RenderBudget.EnsureWithinBudget(1, 1200));

        Assert.Null(exception);
    }
}
