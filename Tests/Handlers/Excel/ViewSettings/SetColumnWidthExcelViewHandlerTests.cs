using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetColumnWidthExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetColumnWidthExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetColumnWidth()
    {
        Assert.Equal("set_column_width", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsColumnWidth()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "width", 20.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("width set to 20", result.ToLower());
        Assert.Equal(20.0, workbook.Worksheets[0].Cells.GetColumnWidth(0));
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsColumnWidthOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "columnIndex", 2 },
            { "width", 15.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("column 2", result.ToLower());
        Assert.Equal(15.0, workbook.Worksheets[1].Cells.GetColumnWidth(2));
    }

    #endregion

    #region Boundary Condition Tests

    [Fact]
    public void Execute_WithoutColumnIndex_UsesDefaultColumnIndex()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", 20.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("column 0", result.ToLower());
        Assert.Contains("width set to 20", result.ToLower());
    }

    [Fact]
    public void Execute_WithoutWidth_UsesDefaultWidth()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("width set to", result.ToLower());
    }

    [Fact]
    public void Execute_WithNegativeColumnIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", -1 },
            { "width", 20.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithZeroWidth_SetsZeroWidth()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "width", 0.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("width set to", result.ToLower());
    }

    [Fact]
    public void Execute_WithNegativeWidth_ThrowsException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "width", -10.0 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithSmallPositiveWidth_SetsWidth()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "width", 0.1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("width set to", result.ToLower());
    }

    [Fact]
    public void Execute_WithLargeWidth_SetsWidth()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "width", 255.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("width set to 255", result.ToLower());
    }

    [Fact]
    public void Execute_WithColumnIndexZero_SetsFirstColumnWidth()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "width", 25.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("column 0", result.ToLower());
    }

    #endregion
}
