using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Filter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Filter;

public class GetExcelFilterStatusHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelFilterStatusHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStatus()
    {
        Assert.Equal("get_status", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_ReturnsCorrectSheetInfo()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("FilteredSheet");
        workbook.Worksheets[1].Cells["A1"].Value = "Data";
        workbook.Worksheets[1].AutoFilter.Range = "A1:C20";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal("FilteredSheet", json.RootElement.GetProperty("worksheetName").GetString());
        Assert.True(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
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
        var workbook = CreateWorkbookWithFilter();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithFilter()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["A2"].Value = "Item1";
        sheet.Cells["B2"].Value = 10;
        sheet.AutoFilter.Range = "A1:B10";
        return workbook;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsFilterStatus()
    {
        var workbook = CreateWorkbookWithFilter();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("isFilterEnabled", out _));
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithFilter_ReturnsEnabled()
    {
        var workbook = CreateWorkbookWithFilter();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Equal("A1:B10", json.RootElement.GetProperty("filterRange").GetString());
    }

    [Fact]
    public void Execute_WithoutFilter_ReturnsDisabled()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Equal(JsonValueKind.Null, json.RootElement.GetProperty("filterRange").ValueKind);
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal("Sheet1", json.RootElement.GetProperty("worksheetName").GetString());
    }

    [Fact]
    public void Execute_ReturnsStatusMessage()
    {
        var workbook = CreateWorkbookWithFilter();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("status", out var status));
        Assert.Contains("Auto filter enabled", status.GetString());
    }

    #endregion
}
