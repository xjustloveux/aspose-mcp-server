using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Results.Excel.Range;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class GetExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Include Format

    [Fact]
    public void Execute_WithIncludeFormat_ReturnsFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "includeFormat", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.NotNull(result.Items[0].Format);
        Assert.NotNull(result.Items[0].Format!.FontName);
        Assert.True(result.Items[0].Format!.FontSize > 0);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyWorkbook()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Original" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Original", workbook.Worksheets[0].Cells["A1"].Value);
        AssertNotModified(context);
    }

    #endregion

    #region Empty Cells

    [Fact]
    public void Execute_WithEmptyCells_ReturnsValue()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.NotNull(result.Items[0].Value);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsRangeData()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A", "B" },
            { "C", "D" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.Equal("A1:B2", result.Range);
        Assert.Equal(4, result.Count);
        AssertNotModified(context);
    }

    [Theory]
    [InlineData("A1:A1", 1, 1)]
    [InlineData("A1:B2", 2, 2)]
    [InlineData("A1:C3", 3, 3)]
    public void Execute_ReturnsCorrectDimensions(string range, int expectedRows, int expectedCols)
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "1", "2", "3" },
            { "4", "5", "6" },
            { "7", "8", "9" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", range }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.Equal(expectedRows, result.RowCount);
        Assert.Equal(expectedCols, result.ColumnCount);
    }

    #endregion

    #region Items Array

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Value1", "Value2" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.NotNull(result.Items);
        Assert.Equal(2, result.Items.Count);
    }

    [Fact]
    public void Execute_ItemsContainCellReference()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.Equal("A1", result.Items[0].Cell);
    }

    [Fact]
    public void Execute_ItemsContainValue()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "TestValue" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.Equal("TestValue", result.Items[0].Value);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_GetsFromSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = "Sheet2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.Contains("Sheet2", result.Items[0].Value);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "FirstSheet";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = "SecondSheet";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.Contains("FirstSheet", result.Items[0].Value);
    }

    #endregion

    #region Include Formulas

    [Fact]
    public void Execute_WithIncludeFormulas_ReturnsFormula()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Formula = "=A1*2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1" },
            { "includeFormulas", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.NotNull(result.Items[0].Formula);
        Assert.Contains("A1", result.Items[0].Formula);
    }

    [Fact]
    public void Execute_WithoutIncludeFormulas_FormulaIsOmitted()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Formula = "=A1*2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1" },
            { "includeFormulas", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRangeResult>(res);

        Assert.Null(result.Items[0].Formula);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("InvalidRange")]
    [InlineData("")]
    public void Execute_WithInvalidRange_ThrowsException(string invalidRange)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", invalidRange }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
