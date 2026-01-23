using AsposeMcpServer.Handlers.Excel.Cell;
using AsposeMcpServer.Results.Excel.Cell;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Cell;

public class GetExcelCellHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelCellHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Theory]
    [InlineData(0, "Value0")]
    [InlineData(1, "Value1")]
    [InlineData(2, "Value2")]
    public void Execute_WithSheetIndex_GetsFromCorrectSheet(int sheetIndex, string expectedValue)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets[0].Cells["A1"].Value = "Value0";
        workbook.Worksheets[1].Cells["A1"].Value = "Value1";
        workbook.Worksheets[2].Cells["A1"].Value = "Value2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "sheetIndex", sheetIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Equal(expectedValue, result.Value);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Original" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
        Assert.Equal("Original", workbook.Worksheets[0].Cells["A1"].Value);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsStringValue()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test Value" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Equal("Test Value", result.Value);
        AssertNotModified(context);
    }

    [Theory]
    [InlineData("A1", "Value A1")]
    [InlineData("B2", "Value B2")]
    [InlineData("C3", "Value C3")]
    public void Execute_ReturnsValueFromCorrectCell(string cell, string expectedValue)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells[cell].Value = expectedValue;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", cell }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Equal(expectedValue, result.Value);
        AssertNotModified(context);
    }

    #endregion

    #region Data Types

    [Fact]
    public void Execute_ReturnsNumericValue()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { 123.45 } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Contains("123.45", result.Value);
    }

    [Fact]
    public void Execute_ReturnsEmptyCellValue()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Contains("empty", result.Value, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsDateValue()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = new DateTime(2024, 1, 15);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.NotNull(result.Value);
    }

    [Fact]
    public void Execute_ReturnsBooleanValue()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = true;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Contains("true", result.Value, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Formula Handling

    [Fact]
    public void Execute_WithFormula_ReturnsFormulaAndValue()
    {
        var workbook = CreateEmptyWorkbook();
        var ws = workbook.Worksheets[0];
        ws.Cells["A1"].Value = 10;
        ws.Cells["B1"].Value = 20;
        ws.Cells["C1"].Formula = "=A1+B1";
        workbook.CalculateFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "includeFormula", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Contains("30", result.Value);
        Assert.Contains("A1+B1", result.Formula);
    }

    [Fact]
    public void Execute_WithCalculateFormula_CalculatesBeforeReturning()
    {
        var workbook = CreateEmptyWorkbook();
        var ws = workbook.Worksheets[0];
        ws.Cells["A1"].Value = 5;
        ws.Cells["B1"].Value = 10;
        ws.Cells["C1"].Formula = "=A1*B1";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "calculateFormula", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Contains("50", result.Value);
    }

    [Fact]
    public void Execute_WithIncludeFormulaFalse_DoesNotIncludeFormula()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Formula = "=1+1";
        workbook.CalculateFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "includeFormula", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Null(result.Formula);
    }

    #endregion

    #region Format Information

    [Fact]
    public void Execute_WithIncludeFormat_ReturnsFormatInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Formatted";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        style.Font.Size = 14;
        cell.SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "includeFormat", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.NotNull(result.Format);
        Assert.True(result.Format!.Bold);
        Assert.Equal(14, result.Format!.FontSize);
    }

    [Fact]
    public void Execute_WithIncludeFormatFalse_DoesNotIncludeFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "includeFormat", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.Null(result.Format);
    }

    #endregion

    #region JSON Response Structure

    [Fact]
    public void Execute_ReturnsValidJsonStructure()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCellResult>(res);

        Assert.NotNull(result.Cell);
        Assert.NotNull(result.Value);
        Assert.NotNull(result.ValueType);
    }

    [Fact]
    public void Execute_ReturnsCellAddressInResponse()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B5" }
        });

        workbook.Worksheets[0].Cells["B5"].Value = "Test";
        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetCellResult>(res);

        Assert.Equal("B5", result.Cell);
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
