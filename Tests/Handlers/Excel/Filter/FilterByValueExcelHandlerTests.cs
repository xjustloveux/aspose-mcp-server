using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Filter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Filter;

public class FilterByValueExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly FilterByValueExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Filter()
    {
        Assert.Equal("filter", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_AppliesFilterToCorrectSheet()
    {
        var workbook = CreateWorkbookWithData();
        workbook.Worksheets.Add("Sheet2");
        FillData(workbook.Worksheets[1]);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "range", "A1:B10" },
            { "criteria", "Item1" }
        });

        _handler.Execute(context, parameters);

        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
        Assert.Equal("A1:B10", workbook.Worksheets[1].AutoFilter.Range);
    }

    #endregion

    #region Basic Filter Operations

    [Fact]
    public void Execute_AppliesFilter()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B10" },
            { "criteria", "Item1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Filter applied", result.Message);
        Assert.Contains("Item1", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsColumnIndex()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B10" },
            { "criteria", "Test" },
            { "columnIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("column 1", result.Message);
    }

    [Fact]
    public void Execute_WithFilterOperator_ReturnsOperatorInMessage()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B10" },
            { "criteria", "100" },
            { "filterOperator", "GreaterThan" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("GreaterThan", result.Message);
    }

    [Theory]
    [InlineData("Equal")]
    [InlineData("NotEqual")]
    [InlineData("GreaterThan")]
    [InlineData("LessThan")]
    public void Execute_WithVariousOperators_Succeeds(string filterOperator)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B10" },
            { "criteria", "10" },
            { "filterOperator", filterOperator }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Filter applied", result.Message);
        Assert.Contains(filterOperator, result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "criteria", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutCriteria_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B10" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("criteria", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "range", "A1:B10" },
            { "criteria", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        FillData(workbook.Worksheets[0]);
        return workbook;
    }

    private static void FillData(Worksheet sheet)
    {
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["A2"].Value = "Item1";
        sheet.Cells["B2"].Value = 10;
        sheet.Cells["A3"].Value = "Item2";
        sheet.Cells["B3"].Value = 20;
        sheet.Cells["A4"].Value = "Item3";
        sheet.Cells["B4"].Value = 30;
    }

    #endregion
}
