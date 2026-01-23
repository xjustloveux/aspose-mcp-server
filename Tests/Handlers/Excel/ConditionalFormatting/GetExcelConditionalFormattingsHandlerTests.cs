using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ConditionalFormatting;
using AsposeMcpServer.Results.Excel.ConditionalFormatting;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetConditionalFormattingsResult>(res);

        Assert.Equal(1, result.SheetIndex);
        Assert.Equal(1, result.Count);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetConditionalFormattingsResult>(res);

        Assert.True(result.Count >= 0);
        Assert.True(result.SheetIndex >= 0);
        Assert.NotNull(result.WorksheetName);
        Assert.NotNull(result.Items);
    }

    [Fact]
    public void Execute_WithNoFormattings_ReturnsEmptyResult()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetConditionalFormattingsResult>(res);

        Assert.Equal(0, result.Count);
        Assert.NotNull(result.Message);
        Assert.Contains("no conditional", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateWorkbookWithMultipleConditionalFormattings(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetConditionalFormattingsResult>(res);

        Assert.Equal(3, result.Count);
        Assert.Equal(3, result.Items.Count);
    }

    [Fact]
    public void Execute_ReturnsConditionDetails()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetConditionalFormattingsResult>(res);

        var item = result.Items[0];
        Assert.True(item.Index >= 0);
        Assert.NotNull(item.Areas);
        Assert.True(item.ConditionsCount >= 0);
        Assert.NotNull(item.Conditions);
    }

    [Fact]
    public void Execute_ReturnsOperatorAndFormula()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetConditionalFormattingsResult>(res);

        var condition = result.Items[0].Conditions[0];
        Assert.NotNull(condition.OperatorType);
        Assert.NotNull(condition.Formula1);
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
