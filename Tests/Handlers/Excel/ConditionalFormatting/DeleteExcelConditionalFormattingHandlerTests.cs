using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ConditionalFormatting;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ConditionalFormatting;

public class DeleteExcelConditionalFormattingHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelConditionalFormattingHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_DeletesFromCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        AddConditionalFormatting(workbook.Worksheets[0], "A1:A5", 10);
        AddConditionalFormatting(workbook.Worksheets[1], "B1:B5", 20);
        AddConditionalFormatting(workbook.Worksheets[2], "C1:C5", 30);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "conditionalFormattingIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].ConditionalFormattings);
        Assert.Empty(workbook.Worksheets[1].ConditionalFormattings);
        Assert.Single(workbook.Worksheets[2].ConditionalFormattings);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesConditionalFormatting()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        Assert.Single(workbook.Worksheets[0].ConditionalFormattings);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Empty(workbook.Worksheets[0].ConditionalFormattings);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var workbook = CreateWorkbookWithMultipleConditionalFormattings(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("remaining: 2", result.ToLower());
        Assert.Equal(2, workbook.Worksheets[0].ConditionalFormattings.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message.ToLower());
    }

    [Fact]
    public void Execute_WithNegativeIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "conditionalFormattingIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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
