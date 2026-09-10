using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Handlers.Excel.FreezePanes;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Excel.FreezePanes;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel;

/// <summary>
///     Guards RB-22, RB-23 and NEW-SORT-01. Sorting copied only <c>Cell.Value</c>, so a formula
///     became the constant it happened to evaluate to and every other per-cell attribute stayed at
///     the old coordinates. The two freeze-panes tools disagreed by one row about what the same
///     request meant, and the read path subtracted one to hide it.
/// </summary>
public class SortAndFreezeContractTests : TestBase
{
    /// <summary>Runs the sort operation.</summary>
    /// <param name="workbook">Workbook to sort.</param>
    /// <param name="range">Range in A1 notation.</param>
    /// <param name="sortColumn">Column index relative to the range.</param>
    /// <param name="ascending">Sort direction.</param>
    /// <param name="hasHeader">Whether the first row is a header.</param>
    private static void Sort(Workbook workbook, string range, int sortColumn, bool ascending, bool hasHeader)
    {
        var handler = new SortDataHandler();
        var context = new OperationContext<Workbook> { Document = workbook };
        var parameters = new OperationParameters();
        parameters.Set("range", range);
        parameters.Set("sortColumn", sortColumn);
        parameters.Set("ascending", ascending);
        parameters.Set("hasHeader", hasHeader);
        handler.Execute(context, parameters);
    }

    [Fact]
    public void Sort_ShouldMoveFormulasRatherThanFlattenThem()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;
        cells["A1"].PutValue(3);
        cells["A2"].PutValue(1);
        cells["A3"].PutValue(2);
        cells["B1"].Formula = "=A1*10";
        cells["B2"].Formula = "=A2*10";
        cells["B3"].Formula = "=A3*10";
        workbook.CalculateFormula();

        Sort(workbook, "A1:B3", 0, true, false);

        Assert.All(["B1", "B2", "B3"],
            address => Assert.True(cells[address].IsFormula, $"{address} lost its formula"));
    }

    [Fact]
    public void Sort_ShouldMoveCellStylesWithTheirRow()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;
        cells["A1"].PutValue("b");
        cells["A2"].PutValue("a");

        var marked = workbook.CreateStyle();
        marked.Font.IsBold = true;
        cells["A1"].SetStyle(marked);

        Sort(workbook, "A1:A2", 0, true, false);

        Assert.Equal("a", cells["A1"].StringValue);
        Assert.Equal("b", cells["A2"].StringValue);
        Assert.True(cells["A2"].GetStyle().Font.IsBold);
        Assert.False(cells["A1"].GetStyle().Font.IsBold);
    }

    [Fact]
    public void Sort_HeaderRow_ShouldStayInPlace()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;
        cells["A1"].PutValue("header");
        cells["A2"].PutValue("b");
        cells["A3"].PutValue("a");

        Sort(workbook, "A1:A3", 0, true, true);

        Assert.Equal("header", cells["A1"].StringValue);
        Assert.Equal("a", cells["A2"].StringValue);
        Assert.Equal("b", cells["A3"].StringValue);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(5)]
    public void Sort_ColumnOutsideTheRange_ShouldBeRejected(int sortColumn)
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("x");

        Assert.Throws<ArgumentException>(() => Sort(workbook, "A1:B2", sortColumn, true, false));
    }

    [Fact]
    public void FreezePanes_BothToolsShouldProduceTheSameSplit()
    {
        var viaFreezeTool = new Workbook();
        var freezeHandler = new FreezeExcelPanesHandler();
        var freezeParams = new OperationParameters();
        freezeParams.Set("row", 2);
        freezeParams.Set("column", 1);
        freezeHandler.Execute(new OperationContext<Workbook> { Document = viaFreezeTool }, freezeParams);

        var viaViewTool = new Workbook();
        var viewHandler = new FreezePanesExcelViewHandler();
        var viewParams = new OperationParameters();
        viewParams.Set("freezeRow", 2);
        viewParams.Set("freezeColumn", 1);
        viewHandler.Execute(new OperationContext<Workbook> { Document = viaViewTool }, viewParams);

        viaFreezeTool.Worksheets[0].GetFreezedPanes(out var rowA, out var colA, out var rowsA, out var colsA);
        viaViewTool.Worksheets[0].GetFreezedPanes(out var rowB, out var colB, out var rowsB, out var colsB);

        Assert.Equal(rowB, rowA);
        Assert.Equal(colB, colA);
        Assert.Equal(rowsB, rowsA);
        Assert.Equal(colsB, colsA);
    }

    [Fact]
    public void FreezePanes_ReadBackShouldMatchTheRequest()
    {
        var workbook = new Workbook();
        var freezeParams = new OperationParameters();
        freezeParams.Set("row", 2);
        freezeParams.Set("column", 1);
        new FreezeExcelPanesHandler().Execute(new OperationContext<Workbook> { Document = workbook }, freezeParams);

        var getParams = new OperationParameters();
        var result = new GetExcelFreezePanesHandler()
            .Execute(new OperationContext<Workbook> { Document = workbook }, getParams);

        var typed = Assert.IsType<GetFreezePanesResult>(result);
        Assert.True(typed.IsFrozen);
        Assert.Equal(2, typed.FrozenRow);
        Assert.Equal(1, typed.FrozenColumn);
    }
}
