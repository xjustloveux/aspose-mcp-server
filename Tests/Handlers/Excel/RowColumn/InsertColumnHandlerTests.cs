using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.RowColumn;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.RowColumn;

public class InsertColumnHandlerTests : ExcelHandlerTestBase
{
    private readonly InsertColumnHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertColumn()
    {
        Assert.Equal("insert_column", _handler.Operation);
    }

    #endregion

    #region Count Parameter

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    public void Execute_WithVariousCounts_InsertsColumns(int count)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", count }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"{count} column(s)", result.Message);
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Inserted", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsColumnIndex()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("column 2", result.Message);
    }

    [Fact]
    public void Execute_ReturnsCount()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3 column(s)", result.Message);
    }

    [Fact]
    public void Execute_DefaultCount_InsertsOneColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("1 column(s)", result.Message);
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_InsertsOnCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Inserted", result.Message);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_UsesFirstSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Inserted", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeColumnIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid column index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", -1 },
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithZeroOrNegativeCount_AcceptsValue(int count)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", count }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"{count} column(s)", result.Message);
    }

    [Fact]
    public void Execute_WithColumnIndexZero_InsertsAtFirstColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Inserted", result.Message);
        Assert.Contains("column 0", result.Message);
    }

    [Fact]
    public void Execute_WithLargeCount_InsertsMultipleColumns()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", 100 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("100 column(s)", result.Message);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Col A");
        sheet.Cells["B1"].PutValue("Col B");
        sheet.Cells["C1"].PutValue("Col C");
        return workbook;
    }

    private new static Workbook CreateWorkbookWithSheets(int sheetCount)
    {
        var workbook = new Workbook();
        for (var i = 1; i < sheetCount; i++)
            workbook.Worksheets.Add();
        return workbook;
    }

    #endregion
}
