using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

public class SortDataHandlerTests : ExcelHandlerTestBase
{
    private readonly SortDataHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Sort()
    {
        Assert.Equal("sort", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Sort Operations

    [Fact]
    public void Execute_SortsDataAscending()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(3);
        workbook.Worksheets[0].Cells["A2"].PutValue(1);
        workbook.Worksheets[0].Cells["A3"].PutValue(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sorted", result.ToLower());
        Assert.Contains("ascending", result.ToLower());
        Assert.Equal(1, workbook.Worksheets[0].Cells["A1"].IntValue);
        Assert.Equal(2, workbook.Worksheets[0].Cells["A2"].IntValue);
        Assert.Equal(3, workbook.Worksheets[0].Cells["A3"].IntValue);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SortsDataDescending()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(1);
        workbook.Worksheets[0].Cells["A2"].PutValue(3);
        workbook.Worksheets[0].Cells["A3"].PutValue(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A3" },
            { "ascending", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("descending", result.ToLower());
        Assert.Equal(3, workbook.Worksheets[0].Cells["A1"].IntValue);
        Assert.Equal(2, workbook.Worksheets[0].Cells["A2"].IntValue);
        Assert.Equal(1, workbook.Worksheets[0].Cells["A3"].IntValue);
    }

    [Fact]
    public void Execute_WithSortColumn_SortsBySpecificColumn()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("C");
        workbook.Worksheets[0].Cells["B1"].PutValue(3);
        workbook.Worksheets[0].Cells["A2"].PutValue("A");
        workbook.Worksheets[0].Cells["B2"].PutValue(1);
        workbook.Worksheets[0].Cells["A3"].PutValue("B");
        workbook.Worksheets[0].Cells["B3"].PutValue(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B3" },
            { "sortColumn", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("column 1", result.ToLower());
        Assert.Equal(1, workbook.Worksheets[0].Cells["B1"].IntValue);
        Assert.Equal(2, workbook.Worksheets[0].Cells["B2"].IntValue);
        Assert.Equal(3, workbook.Worksheets[0].Cells["B3"].IntValue);
    }

    [Fact]
    public void Execute_WithHeader_PreservesHeader()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Header");
        workbook.Worksheets[0].Cells["A2"].PutValue(3);
        workbook.Worksheets[0].Cells["A3"].PutValue(1);
        workbook.Worksheets[0].Cells["A4"].PutValue(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A4" },
            { "hasHeader", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Header", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal(1, workbook.Worksheets[0].Cells["A2"].IntValue);
        Assert.Equal(2, workbook.Worksheets[0].Cells["A3"].IntValue);
        Assert.Equal(3, workbook.Worksheets[0].Cells["A4"].IntValue);
    }

    [Fact]
    public void Execute_WithSheetIndex_SortsInSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        workbook.Worksheets[1].Cells["A1"].PutValue(2);
        workbook.Worksheets[1].Cells["A2"].PutValue(1);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A2" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1, workbook.Worksheets[1].Cells["A1"].IntValue);
        Assert.Equal(2, workbook.Worksheets[1].Cells["A2"].IntValue);
    }

    #endregion
}
