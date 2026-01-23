using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.MergeCells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.MergeCells;

public class UnmergeExcelCellsHandlerTests : ExcelHandlerTestBase
{
    private readonly UnmergeExcelCellsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Unmerge()
    {
        Assert.Equal("unmerge", _handler.Operation);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_UnmergesOnCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[0].Cells.CreateRange("A1:B2").Merge();
        workbook.Worksheets[1].Cells.CreateRange("A1:B2").Merge();
        workbook.Worksheets[2].Cells.CreateRange("A1:B2").Merge();

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);
        Assert.Empty(workbook.Worksheets[1].Cells.MergedCells);
        Assert.Single(workbook.Worksheets[2].Cells.MergedCells);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithMergedCells()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells.CreateRange("A1:C3").Merge();
        return workbook;
    }

    #endregion

    #region Basic Unmerge Operations

    [Fact]
    public void Execute_UnmergesCells()
    {
        var workbook = CreateWorkbookWithMergedCells();
        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("unmerged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
        AssertModified(context);
    }

    [Fact]
    public void Execute_UnmergesOnlySpecifiedRange()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells.CreateRange("A1:B2").Merge();
        workbook.Worksheets[0].Cells.CreateRange("D1:E2").Merge();
        Assert.Equal(2, workbook.Worksheets[0].Cells.MergedCells.Count);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Execute_WithNotMergedRange_StillSucceeds()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("unmerged", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
