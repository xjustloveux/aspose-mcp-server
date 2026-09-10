using System.Text.Json;
using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

/// <summary>
///     Covers LOW-10. Reading through the <c>Cells[row, column]</c> indexer instantiates a cell
///     that did not exist, so asking for statistics over a large mostly-empty range turned every
///     blank position into a real cell. In session mode those cells stay in the workbook and are
///     written out on the next save, growing the file for a read-only query.
/// </summary>
public class StatisticsCellMaterializationTests : ExcelHandlerTestBase
{
    private readonly GetStatisticsHandler _handler = new();

    [Fact]
    public void Execute_WithLargeSparseRange_ShouldNotCreateCells()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["A3"].PutValue(3);
        var before = sheet.Cells.Count;

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "range", "A1:Z1000" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(before, sheet.Cells.Count);
    }

    [Fact]
    public void Execute_WithLargeSparseRange_ShouldStillCountTheValues()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["C3"].PutValue("text");

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "range", "A1:Z1000" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("15", json);
    }

    [Fact]
    public void GetFormulas_OverASparseSheet_ShouldNotCreateCells()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["Z400"].Formula = "=SUM(A1:A2)";
        var before = sheet.Cells.Count;

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var result = new GetFormulasHandler().Execute(context, parameters);

        Assert.Equal(before, sheet.Cells.Count);
        Assert.Contains("SUM(A1:A2)", JsonSerializer.Serialize(result));
    }
}
