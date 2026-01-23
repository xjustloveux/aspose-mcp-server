using AsposeMcpServer.Handlers.Excel.Cell;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Cell;

public class ClearExcelCellHandlerTests : ExcelHandlerTestBase
{
    private readonly ClearExcelCellHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Clear()
    {
        Assert.Equal("clear", _handler.Operation);
    }

    #endregion

    #region Clear Content Only

    [Fact]
    public void Execute_WithClearContentTrue_ClearsOnlyContent()
    {
        var workbook = CreateEmptyWorkbook();
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "clearContent", true },
            { "clearFormat", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.True(workbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
    }

    #endregion

    #region Clear Format Only

    [Fact]
    public void Execute_WithClearFormatTrue_ClearsOnlyFormat()
    {
        var workbook = CreateEmptyWorkbook();
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "clearContent", false },
            { "clearFormat", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.False(workbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
    }

    #endregion

    #region Clear Both

    [Fact]
    public void Execute_WithBothTrue_ClearsContentAndFormat()
    {
        var workbook = CreateEmptyWorkbook();
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "clearContent", true },
            { "clearFormat", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.False(workbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
    }

    #endregion

    #region Sheet Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithSheetIndex_ClearsFromCorrectSheet(int sheetIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets[sheetIndex].Cells["A1"].Value = "Test";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "sheetIndex", sheetIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[sheetIndex].Cells["A1"].StringValue);
    }

    #endregion

    #region Preserve Other Cells

    [Fact]
    public void Execute_PreservesOtherCells()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A1", "B1" },
            { "A2", "B2" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("B1", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.Equal("A2", workbook.Worksheets[0].Cells["A2"].Value);
        Assert.Equal("B2", workbook.Worksheets[0].Cells["B2"].Value);
    }

    #endregion

    #region Basic Clear Operations

    [Fact]
    public void Execute_ClearsCellContent()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Hello" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("B2")]
    [InlineData("Z10")]
    public void Execute_ClearsVariousCells(string cell)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells[cell].Value = "Test Value";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", cell }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("", workbook.Worksheets[0].Cells[cell].StringValue);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("InvalidCell")]
    [InlineData("1A")]
    public void Execute_WithInvalidCellAddress_ThrowsException(string invalidCell)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", invalidCell }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
