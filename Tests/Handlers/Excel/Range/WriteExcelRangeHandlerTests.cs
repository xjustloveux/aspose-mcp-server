using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class WriteExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly WriteExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Write()
    {
        Assert.Equal("write", _handler.Operation);
    }

    #endregion

    #region Large Range

    [Fact]
    public void Execute_WritesLargeRange()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", "A1" },
            {
                "data",
                "[[\"1\",\"2\",\"3\",\"4\",\"5\"],[\"6\",\"7\",\"8\",\"9\",\"10\"],[\"11\",\"12\",\"13\",\"14\",\"15\"]]"
            }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1.0, workbook.Worksheets[0].Cells["A1"].DoubleValue);
        Assert.Equal(15.0, workbook.Worksheets[0].Cells["E3"].DoubleValue);
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
            { "startCell", "B3" },
            { "data", "[[\"Test\"]]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B3", result);
        Assert.Contains("written", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Write Operations

    [Fact]
    public void Execute_Writes2DArrayData()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", "A1" },
            { "data", "[[\"A\",\"B\"],[\"C\",\"D\"]]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1", result);
        Assert.Equal("A", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("B", workbook.Worksheets[0].Cells["B1"].StringValue);
        Assert.Equal("C", workbook.Worksheets[0].Cells["A2"].StringValue);
        Assert.Equal("D", workbook.Worksheets[0].Cells["B2"].StringValue);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("B2")]
    [InlineData("C5")]
    public void Execute_WritesToVariousStartCells(string startCell)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", startCell },
            { "data", "[[\"Test\"]]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains(startCell, result);
        Assert.Equal("Test", workbook.Worksheets[0].Cells[startCell].StringValue);
        AssertModified(context);
    }

    #endregion

    #region Data Types

    [Fact]
    public void Execute_WritesNumericData()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", "A1" },
            { "data", "[[\"1\",\"2\",\"3\"],[\"4\",\"5\",\"6\"]]" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1.0, workbook.Worksheets[0].Cells["A1"].DoubleValue);
        Assert.Equal(6.0, workbook.Worksheets[0].Cells["C2"].DoubleValue);
    }

    [Fact]
    public void Execute_WritesMixedData()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", "A1" },
            { "data", "[[\"Name\",\"Age\"],[\"John\",\"30\"]]" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Name", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Age", workbook.Worksheets[0].Cells["B1"].StringValue);
        Assert.Equal("John", workbook.Worksheets[0].Cells["A2"].StringValue);
        Assert.Equal(30.0, workbook.Worksheets[0].Cells["B2"].DoubleValue);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_WritesToSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", "A1" },
            { "data", "[[\"Sheet2 Data\"]]" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Sheet2 Data", workbook.Worksheets[1].Cells["A1"].StringValue);
        Assert.Null(workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_WritesToFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", "A1" },
            { "data", "[[\"First Sheet\"]]" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("First Sheet", workbook.Worksheets[0].Cells["A1"].StringValue);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStartCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", "[[\"Test\"]]" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startCell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutData_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCell", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("data", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
