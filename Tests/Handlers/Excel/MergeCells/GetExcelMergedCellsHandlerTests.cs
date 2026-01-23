using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.MergeCells;
using AsposeMcpServer.Results.Excel.MergeCells;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.MergeCells;

public class GetExcelMergedCellsHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelMergedCellsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_ReturnsCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[1].Cells.CreateRange("A1:B2").Merge();
        workbook.Worksheets[1].Cells.CreateRange("D1:E2").Merge();

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMergedCellsResult>(res);

        Assert.Equal(2, result.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var workbook = CreateWorkbookWithMergedCells();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsMergedCellsInfo()
    {
        var workbook = CreateWorkbookWithMergedCells();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMergedCellsResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.WorksheetName);
        Assert.NotNull(result.Items);
    }

    [Fact]
    public void Execute_WithNoMergedCells_ReturnsEmptyResult()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMergedCellsResult>(res);

        Assert.Equal(0, result.Count);
        Assert.NotNull(result.Message);
        Assert.Contains("no merged", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateWorkbookWithMultipleMergedCells(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMergedCellsResult>(res);

        Assert.Equal(3, result.Count);
        Assert.Equal(3, result.Items.Count);
    }

    [Fact]
    public void Execute_ReturnsMergedCellDetails()
    {
        var workbook = CreateWorkbookWithMergedCells();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMergedCellsResult>(res);

        var item = result.Items[0];
        Assert.True(item.Index >= 0);
        Assert.NotNull(item.Range);
        Assert.NotNull(item.StartCell);
        Assert.NotNull(item.EndCell);
        Assert.True(item.RowCount > 0);
        Assert.True(item.ColumnCount > 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectRangeInfo()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells.CreateRange("B2:D4").Merge();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMergedCellsResult>(res);

        var item = result.Items[0];
        Assert.Equal("B2", item.StartCell);
        Assert.Equal("D4", item.EndCell);
        Assert.Equal(3, item.RowCount);
        Assert.Equal(3, item.ColumnCount);
    }

    [Fact]
    public void Execute_ReturnsCellValue()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Value";
        workbook.Worksheets[0].Cells.CreateRange("A1:B2").Merge();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMergedCellsResult>(res);

        Assert.Equal("Test Value", result.Items[0].Value);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithMergedCells()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells.CreateRange("A1:C3").Merge();
        return workbook;
    }

    private static Workbook CreateWorkbookWithMultipleMergedCells(int count)
    {
        var workbook = new Workbook();
        for (var i = 0; i < count; i++)
        {
            var startRow = i * 5;
            workbook.Worksheets[0].Cells.CreateRange(startRow, 0, 2, 2).Merge();
        }

        return workbook;
    }

    #endregion
}
