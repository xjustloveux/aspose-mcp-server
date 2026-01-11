using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.RowColumn;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.RowColumn;

public class DeleteColumnHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteColumnHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteColumn()
    {
        Assert.Equal("delete_column", _handler.Operation);
    }

    #endregion

    #region Count Parameter

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_WithVariousCounts_DeletesColumns(int count)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", count }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"{count} column(s)", result);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("column 2", result);
    }

    [Fact]
    public void Execute_ReturnsCount()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2 column(s)", result);
    }

    [Fact]
    public void Execute_DefaultCount_DeletesOneColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1 column(s)", result);
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_DeletesFromCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[1].Cells["A1"].PutValue("Data");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "columnIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted", result);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_UsesFirstSheet()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"{count} column(s)", result);
    }

    [Fact]
    public void Execute_WithColumnIndexZero_DeletesFirstColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted", result);
        Assert.Contains("column 0", result);
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
