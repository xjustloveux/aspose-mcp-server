using AsposeMcpServer.Handlers.Excel.Cell;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Cell;

public class WriteExcelCellHandlerTests : ExcelHandlerTestBase
{
    private readonly WriteExcelCellHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Write()
    {
        Assert.Equal("write", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsCorrectResultMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B5" },
            { "value", "TestValue" },
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B5", result);
        Assert.Contains("TestValue", result);
        Assert.Contains("sheet", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Write Operations

    [Fact]
    public void Execute_WritesStringValue()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "Hello World" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1", result);
        Assert.Contains("Hello World", result);
        AssertCellValue(workbook, 0, 0, "Hello World");
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1", "Test Value")]
    [InlineData("B2", "Another Value")]
    [InlineData("Z1", "Column Z")]
    [InlineData("AA1", "Column AA")]
    [InlineData("XFD1", "Last Column")]
    public void Execute_WritesToVariousCellAddresses(string cell, string value)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", cell },
            { "value", value }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains(cell, result);
        Assert.Equal(value, workbook.Worksheets[0].Cells[cell].Value);
        AssertModified(context);
    }

    #endregion

    #region Data Types

    [Theory]
    [InlineData("123", 123)]
    [InlineData("123.45", 123.45)]
    [InlineData("-100", -100)]
    [InlineData("0", 0)]
    public void Execute_WritesNumericValues(string input, object expected)
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
        Assert.Equal(Convert.ToDouble(expected), Convert.ToDouble(cellValue));
        AssertModified(context);
    }

    [Theory]
    [InlineData("true")]
    [InlineData("false")]
    [InlineData("TRUE")]
    [InlineData("FALSE")]
    public void Execute_WritesBooleanValues(string value)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", value }
        });

        _handler.Execute(context, parameters);

        var cellValue = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.NotNull(cellValue);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WritesDateValue()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "2024-01-15" }
        });

        _handler.Execute(context, parameters);

        var dateValue = workbook.Worksheets[0].Cells["A1"].DateTimeValue;
        Assert.Equal(new DateTime(2024, 1, 15), dateValue.Date);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Simple text")]
    [InlineData("Text with numbers 12345")]
    [InlineData("Special chars: !@#$%^&*()")]
    [InlineData("Unicode: 中文測試")]
    public void Execute_WritesVariousStringFormats(string value)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", value }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(value, workbook.Worksheets[0].Cells["A1"].StringValue);
        AssertModified(context);
    }

    #endregion

    #region Sheet Index

    [Theory]
    [InlineData(0, "Value0")]
    [InlineData(1, "Value1")]
    [InlineData(2, "Value2")]
    public void Execute_WithSheetIndex_WritesToCorrectSheet(int sheetIndex, string value)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", value },
            { "sheetIndex", sheetIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(value, workbook.Worksheets[sheetIndex].Cells["A1"].Value);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_WritesToFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "First Sheet" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("First Sheet", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Null(workbook.Worksheets[1].Cells["A1"].Value);
    }

    #endregion

    #region Overwrite Behavior

    [Fact]
    public void Execute_OverwritesExistingValue()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Old Value" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "New Value" }
        });

        _handler.Execute(context, parameters);

        AssertCellValue(workbook, 0, 0, "New Value");
        AssertModified(context);
    }

    [Fact]
    public void Execute_OverwritesNumericWithString()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { 123 } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "Text Value" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Text Value", workbook.Worksheets[0].Cells["A1"].StringValue);
    }

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

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "value", "Hello" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutValue_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("value", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyValue_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "value", "" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData("InvalidCell")]
    [InlineData("1A")]
    [InlineData("")]
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
