using AsposeMcpServer.Handlers.Excel.Cell;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Cell;

public class EditExcelCellHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelCellHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Clear Value

    [Fact]
    public void Execute_WithClearValue_ClearsCell()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Value to Clear" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "clearValue", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        AssertModified(context);
    }

    #endregion

    #region Sheet Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithSheetIndex_EditsCorrectSheet(int sheetIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets[sheetIndex].Cells["A1"].Value = "Original";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "Updated" },
            { "sheetIndex", sheetIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Updated", workbook.Worksheets[sheetIndex].Cells["A1"].StringValue);
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
            { "cell", "A1" },
            { "value", "Modified" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Modified", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("B1", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.Equal("A2", workbook.Worksheets[0].Cells["A2"].Value);
        Assert.Equal("B2", workbook.Worksheets[0].Cells["B2"].Value);
    }

    #endregion

    #region Edit with Value

    [Fact]
    public void Execute_WithValue_UpdatesCellValue()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Old Value" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "New Value" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("New Value", workbook.Worksheets[0].Cells["A1"].StringValue);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1", "Test Value")]
    [InlineData("B2", "Another Value")]
    [InlineData("Z10", "Last Value")]
    public void Execute_WithValue_EditsVariousCells(string cell, string value)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells[cell].Value = "Original";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", cell },
            { "value", value }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(value, workbook.Worksheets[0].Cells[cell].StringValue);
        AssertModified(context);
    }

    [Theory]
    [InlineData("123", 123)]
    [InlineData("45.67", 45.67)]
    [InlineData("-100", -100)]
    public void Execute_WithNumericValue_SetsNumber(string input, double expected)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", input }
        });

        _handler.Execute(context, parameters);

        var cellValue = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.Equal(expected, Convert.ToDouble(cellValue));
    }

    #endregion

    #region Edit with Formula

    [Fact]
    public void Execute_WithFormula_SetsFormula()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Value = 20;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "formula", "=A1+B1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("=A1+B1", workbook.Worksheets[0].Cells["C1"].Formula);
        AssertModified(context);
    }

    [Theory]
    [InlineData("=SUM(A1:A10)")]
    [InlineData("=AVERAGE(B1:B5)")]
    [InlineData("=IF(A1>0,\"Yes\",\"No\")")]
    public void Execute_WithVariousFormulas_SetsFormula(string formula)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "formula", formula }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(formula, workbook.Worksheets[0].Cells["A1"].Formula);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "value", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutValueOrFormula_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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
            { "cell", invalidCell },
            { "value", "Test" }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
