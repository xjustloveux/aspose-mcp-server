using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ConditionalFormatting;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ConditionalFormatting;

public class GetExcelConditionalFormattingsHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelConditionalFormattingsHandler _handler = new();

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
        AddConditionalFormatting(workbook.Worksheets[1], "B1:B10", 20);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("sheetIndex").GetInt32());
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
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
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsFormattingInfo()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("sheetIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
        Assert.True(json.RootElement.TryGetProperty("items", out _));
    }

    [Fact]
    public void Execute_WithNoFormattings_ReturnsEmptyResult()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.True(json.RootElement.TryGetProperty("message", out var message));
        Assert.Contains("no conditional", message.GetString()?.ToLower());
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateWorkbookWithMultipleConditionalFormattings(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void Execute_ReturnsConditionDetails()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var item = json.RootElement.GetProperty("items")[0];
        Assert.True(item.TryGetProperty("index", out _));
        Assert.True(item.TryGetProperty("areas", out _));
        Assert.True(item.TryGetProperty("conditionsCount", out _));
        Assert.True(item.TryGetProperty("conditions", out _));
    }

    [Fact]
    public void Execute_ReturnsOperatorAndFormula()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var condition = json.RootElement.GetProperty("items")[0].GetProperty("conditions")[0];
        Assert.True(condition.TryGetProperty("operatorType", out _));
        Assert.True(condition.TryGetProperty("formula1", out _));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithConditionalFormatting()
    {
        var workbook = new Workbook();
        AddConditionalFormatting(workbook.Worksheets[0], "A1:A10", 50);
        return workbook;
    }

    private static Workbook CreateWorkbookWithMultipleConditionalFormattings(int count)
    {
        var workbook = new Workbook();
        for (var i = 0; i < count; i++)
            AddConditionalFormatting(workbook.Worksheets[0], $"A{i * 10 + 1}:A{i * 10 + 10}", i * 10);
        return workbook;
    }

    private static void AddConditionalFormatting(Worksheet worksheet, string range, int value)
    {
        var formatIndex = worksheet.ConditionalFormattings.Add();
        var fcs = worksheet.ConditionalFormattings[formatIndex];
        var cellArea = CellArea.CreateCellArea(range.Split(':')[0], range.Split(':')[1]);
        fcs.AddArea(cellArea);
        var conditionIndex = fcs.AddCondition(FormatConditionType.CellValue);
        fcs[conditionIndex].Operator = OperatorType.GreaterThan;
        fcs[conditionIndex].Formula1 = value.ToString();
    }

    #endregion
}
