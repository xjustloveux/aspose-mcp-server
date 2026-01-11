using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ConditionalFormatting;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ConditionalFormatting;

public class EditExcelConditionalFormattingHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelConditionalFormattingHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Between Condition

    [Fact]
    public void Execute_WithBetweenAndFormula2_SetsBothFormulas()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 },
            { "conditionIndex", 0 },
            { "condition", "Between" },
            { "value", "10" },
            { "formula2", "20" }
        });

        _handler.Execute(context, parameters);

        var fc = workbook.Worksheets[0].ConditionalFormattings[0][0];
        Assert.Equal(OperatorType.Between, fc.Operator);
        Assert.Contains("10", fc.Formula1);
        Assert.Contains("20", fc.Formula2);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_EditsCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        AddConditionalFormatting(workbook.Worksheets[1], "B1:B10", 20);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "conditionalFormattingIndex", 0 },
            { "conditionIndex", 0 },
            { "value", "99" }
        });

        _handler.Execute(context, parameters);

        Assert.Contains("99", workbook.Worksheets[1].ConditionalFormattings[0][0].Formula1);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsConditionalFormatting()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 },
            { "conditionIndex", 0 },
            { "value", "100" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result.ToLower());
        Assert.Contains("100", workbook.Worksheets[0].ConditionalFormattings[0][0].Formula1);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCondition_ChangesOperator()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        Assert.Equal(OperatorType.GreaterThan, workbook.Worksheets[0].ConditionalFormattings[0][0].Operator);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 },
            { "conditionIndex", 0 },
            { "condition", "LessThan" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(OperatorType.LessThan, workbook.Worksheets[0].ConditionalFormattings[0][0].Operator);
    }

    [Fact]
    public void Execute_WithBackgroundColor_ChangesColor()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 },
            { "conditionIndex", 0 },
            { "backgroundColor", "Blue" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("BackgroundColor=Blue", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsChangesDescription()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 },
            { "conditionIndex", 0 },
            { "condition", "Equal" },
            { "value", "75" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Operator=Equal", result);
        Assert.Contains("Value=75", result);
    }

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChangesMessage()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("no changes", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidFormattingIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 99 },
            { "conditionIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message.ToLower());
    }

    [Fact]
    public void Execute_WithInvalidConditionIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithConditionalFormatting();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "conditionalFormattingIndex", 0 },
            { "conditionIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message.ToLower());
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
