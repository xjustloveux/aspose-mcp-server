using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.MergeCells;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
        Assert.True(json.RootElement.TryGetProperty("items", out _));
    }

    [Fact]
    public void Execute_WithNoMergedCells_ReturnsEmptyResult()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.True(json.RootElement.TryGetProperty("message", out var message));
        Assert.Contains("no merged", message.GetString()?.ToLower());
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateWorkbookWithMultipleMergedCells(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void Execute_ReturnsMergedCellDetails()
    {
        var workbook = CreateWorkbookWithMergedCells();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var item = json.RootElement.GetProperty("items")[0];
        Assert.True(item.TryGetProperty("index", out _));
        Assert.True(item.TryGetProperty("range", out _));
        Assert.True(item.TryGetProperty("startCell", out _));
        Assert.True(item.TryGetProperty("endCell", out _));
        Assert.True(item.TryGetProperty("rowCount", out _));
        Assert.True(item.TryGetProperty("columnCount", out _));
    }

    [Fact]
    public void Execute_ReturnsCorrectRangeInfo()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells.CreateRange("B2:D4").Merge();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var item = json.RootElement.GetProperty("items")[0];
        Assert.Equal("B2", item.GetProperty("startCell").GetString());
        Assert.Equal("D4", item.GetProperty("endCell").GetString());
        Assert.Equal(3, item.GetProperty("rowCount").GetInt32());
        Assert.Equal(3, item.GetProperty("columnCount").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsCellValue()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Value";
        workbook.Worksheets[0].Cells.CreateRange("A1:B2").Merge();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var item = json.RootElement.GetProperty("items")[0];
        Assert.Equal("Test Value", item.GetProperty("value").GetString());
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
