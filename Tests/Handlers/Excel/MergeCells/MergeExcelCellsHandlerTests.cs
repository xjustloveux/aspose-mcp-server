using AsposeMcpServer.Handlers.Excel.MergeCells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.MergeCells;

public class MergeExcelCellsHandlerTests : ExcelHandlerTestBase
{
    private readonly MergeExcelCellsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Merge()
    {
        Assert.Equal("merge", _handler.Operation);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_MergesOnCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
        Assert.Single(workbook.Worksheets[1].Cells.MergedCells);
        Assert.Empty(workbook.Worksheets[2].Cells.MergedCells);
    }

    #endregion

    #region Basic Merge Operations

    [Fact]
    public void Execute_MergesCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("A1:C3", result.Message);
        Assert.True(workbook.Worksheets[0].Cells.MergedCells.Count > 0);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsRowAndColumnCount()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B4" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("4 rows", result.Message);
        Assert.Contains("2 columns", result.Message);
    }

    [Fact]
    public void Execute_MergesTwoHorizontalCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Execute_MergesTwoVerticalCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);
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
    public void Execute_WithSingleCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("single cell", ex.Message, StringComparison.OrdinalIgnoreCase);
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
