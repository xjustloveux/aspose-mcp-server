using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ConditionalFormatting;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ConditionalFormatting;

public class AddExcelConditionalFormattingHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelConditionalFormattingHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_AddsToCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" },
            { "condition", "GreaterThan" },
            { "value", "10" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(workbook.Worksheets[0].ConditionalFormattings);
        Assert.Single(workbook.Worksheets[1].ConditionalFormattings);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsConditionalFormatting()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" },
            { "condition", "GreaterThan" },
            { "value", "10" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("A1:B5", result);
        Assert.Single(workbook.Worksheets[0].ConditionalFormattings);
        AssertModified(context);
    }

    [Theory]
    [InlineData("GreaterThan", OperatorType.GreaterThan)]
    [InlineData("LessThan", OperatorType.LessThan)]
    [InlineData("Equal", OperatorType.Equal)]
    public void Execute_WithVariousConditions_SetsCorrectOperator(string condition, OperatorType expectedOperator)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A10" },
            { "condition", condition },
            { "value", "50" }
        });

        _handler.Execute(context, parameters);

        var fcs = workbook.Worksheets[0].ConditionalFormattings[0];
        Assert.Equal(expectedOperator, fcs[0].Operator);
    }

    [Fact]
    public void Execute_WithBetweenCondition_SetsBothFormulas()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A10" },
            { "condition", "Between" },
            { "value", "10" },
            { "formula2", "20" }
        });

        _handler.Execute(context, parameters);

        var fcs = workbook.Worksheets[0].ConditionalFormattings[0];
        Assert.Equal(OperatorType.Between, fcs[0].Operator);
        Assert.Contains("10", fcs[0].Formula1);
        Assert.Contains("20", fcs[0].Formula2);
    }

    [Fact]
    public void Execute_WithBetweenAndCommaValue_ParsesCorrectly()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A10" },
            { "condition", "Between" },
            { "value", "5, 15" }
        });

        _handler.Execute(context, parameters);

        var fcs = workbook.Worksheets[0].ConditionalFormattings[0];
        Assert.Contains("5", fcs[0].Formula1);
        Assert.Contains("15", fcs[0].Formula2);
    }

    [Fact]
    public void Execute_WithBackgroundColor_SetsColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" },
            { "condition", "GreaterThan" },
            { "value", "10" },
            { "backgroundColor", "Red" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "condition", "GreaterThan" },
            { "value", "10" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutCondition_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" },
            { "value", "10" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutValue_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" },
            { "condition", "GreaterThan" }
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
            { "range", "A1:B5" },
            { "condition", "GreaterThan" },
            { "value", "10" },
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
