using System.Text.Json;
using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

/// <summary>
///     Covers A-05. The statistics handler walks the requested range as a nested row/column loop.
///     <c>CheckCell</c> stopped it from materialising cells (LOW-10) but does nothing about the
///     number of iterations, and a full worksheet is 1,048,576 × 16,384 — about 17 billion
///     coordinates. A single read-only-looking request could therefore occupy a core for a very
///     long time, so the work has to be bounded before the loop starts rather than interrupted
///     part-way.
/// </summary>
public class StatisticsScanBudgetTests : ExcelHandlerTestBase
{
    private readonly GetStatisticsHandler _handler = new();

    [Fact]
    public void Execute_WithAnEnormousRange_ShouldBeRefusedBeforeScanning()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(1);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "range", "A1:XFD1048576" }
        });

        var result = _handler.Execute(context, parameters);

        // Reported in the range error field, which is how this handler surfaces an unusable
        // range; the message has to say what is wrong rather than the generic sentinel.
        var json = JsonSerializer.Serialize(result);
        Assert.Contains("too large", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("internal processing error", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithARangeInsideTheBudget_ShouldStillWork()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "range", "A1:Z1000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("30", JsonSerializer.Serialize(result));
    }

    [Fact]
    public void Execute_WithoutARange_ShouldStillWork()
    {
        // The whole-workbook path reports per-sheet totals from Aspose's own metadata and never
        // enters the nested loop, so the budget must not get in its way.
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("data");
        var context = CreateContext(workbook);

        var result = _handler.Execute(context, CreateEmptyParameters());

        Assert.NotNull(result);
    }

    /// <summary>
    ///     A range this server cannot parse fails inside the spreadsheet library, and that
    ///     exception's text was forwarded to the caller because it happened to be an
    ///     <c>ArgumentException</c> (R3-C09). Authored diagnostics are now told apart by type.
    /// </summary>
    /// <param name="range">A range string the library cannot use.</param>
    [Theory]
    [InlineData("not-a-range")]
    [InlineData("A1:")]
    [InlineData("::")]
    public void Execute_WithAnUnusableRange_ShouldReportAComposedMessage(string range)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(1);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "range", range }
        });

        var json = JsonSerializer.Serialize(_handler.Execute(context, parameters));

        Assert.Contains("could not be read", json, StringComparison.Ordinal);
        Assert.DoesNotContain("Exception", json, StringComparison.Ordinal);
        Assert.DoesNotContain("Aspose", json, StringComparison.Ordinal);
    }
}
