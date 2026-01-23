using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class ClearExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly ClearExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Clear()
    {
        Assert.Equal("clear", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_ClearsFromSpecificSheet()
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

        _handler.Execute(context, parameters);

        Assert.Equal("Sheet1", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("", workbook.Worksheets[1].Cells["A1"].StringValue);
    }

    #endregion

    #region Preserve Other Cells

    [Fact]
    public void Execute_PreservesOtherCells()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A1", "B1", "C1" },
            { "A2", "B2", "C2" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("B1", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.Equal("A2", workbook.Worksheets[0].Cells["A2"].Value);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("A1:B2", result.Message);
        Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Clear Operations

    [Fact]
    public void Execute_ClearsRangeContent()
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("", workbook.Worksheets[0].Cells["B2"].StringValue);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("A1:B2")]
    [InlineData("A1:C3")]
    public void Execute_ClearsVariousRanges(string range)
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Clear Options

    [Fact]
    public void Execute_WithClearContentTrue_ClearsContent()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "clearContent", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
    }

    [Fact]
    public void Execute_WithClearFormatTrue_ClearsFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "clearContent", false },
            { "clearFormat", true }
        });

        _handler.Execute(context, parameters);

        Assert.False(workbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
    }

    [Fact]
    public void Execute_WithBothClearOptions_ClearsContentAndFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "clearContent", true },
            { "clearFormat", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.False(workbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
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
