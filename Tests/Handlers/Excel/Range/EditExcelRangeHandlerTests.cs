using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class EditExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_EditsSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = "Sheet2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "data", "[[\"Modified\"]]" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Sheet1", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("Modified", workbook.Worksheets[1].Cells["A1"].StringValue);
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
            { "range", "A1" },
            { "data", "[[\"Modified\"]]" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Modified", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("B1", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.Equal("C1", workbook.Worksheets[0].Cells["C1"].Value);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "data", "[[\"X\"]]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1:B2", result);
        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsRangeData()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Old1", "Old2" },
            { "Old3", "Old4" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "data", "[[\"New1\",\"New2\"],[\"New3\",\"New4\"]]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("New1", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("New4", workbook.Worksheets[0].Cells["B2"].StringValue);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("A1:B2")]
    [InlineData("B2:C3")]
    public void Execute_EditsVariousRanges(string range)
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
            { "range", range },
            { "data", "[[\"X\"]]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Clear Range Option

    [Fact]
    public void Execute_WithClearRangeTrue_ClearsBeforeEdit()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A", "B", "C" },
            { "D", "E", "F" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C2" },
            { "data", "[[\"X\"]]" },
            { "clearRange", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("X", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("", workbook.Worksheets[0].Cells["B1"].StringValue);
        Assert.Equal("", workbook.Worksheets[0].Cells["C2"].StringValue);
    }

    [Fact]
    public void Execute_WithClearRangeFalse_PreservesUnedited()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A", "B" },
            { "C", "D" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "data", "[[\"X\"]]" },
            { "clearRange", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("X", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("B", workbook.Worksheets[0].Cells["B1"].Value);
    }

    #endregion

    #region Data Types

    [Fact]
    public void Execute_EditsWithNumericData()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Old" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "data", "[[\"123\"]]" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(123.0, workbook.Worksheets[0].Cells["A1"].DoubleValue);
    }

    [Fact]
    public void Execute_EditsWithMixedData()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B1" },
            { "data", "[[\"Text\",\"100\"]]" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Text", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal(100.0, workbook.Worksheets[0].Cells["B1"].DoubleValue);
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
            { "data", "[[\"Test\"]]" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutData_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("data", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
