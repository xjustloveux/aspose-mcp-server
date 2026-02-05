using AsposeMcpServer.Handlers.Excel.DataImportExport;
using AsposeMcpServer.Results.Excel.DataImportExport;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataImportExport;

public class ImportArrayExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly ImportArrayExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeImportArray()
    {
        Assert.Equal("import_array", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithSingleRow_ShouldImport()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "arrayData", "A,B,C" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ImportExcelResult>(res);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal("A1", result.StartCell);

        AssertCellValue(workbook, 0, 0, "A");
        AssertCellValue(workbook, 0, 1, "B");
        AssertCellValue(workbook, 0, 2, "C");
    }

    [Fact]
    public void Execute_WithMultipleRows_ShouldImport()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "arrayData", "A,B,C;1,2,3;4,5,6" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ImportExcelResult>(res);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);

        AssertCellValue(workbook, 0, 0, "A");
        AssertCellValue(workbook, 0, 1, "B");
        AssertCellValue(workbook, 0, 2, "C");
        AssertCellValue(workbook, 1, 0, "1");
        AssertCellValue(workbook, 1, 1, "2");
        AssertCellValue(workbook, 1, 2, "3");
        AssertCellValue(workbook, 2, 0, "4");
        AssertCellValue(workbook, 2, 1, "5");
        AssertCellValue(workbook, 2, 2, "6");
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "arrayData", "A,B,C" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMissingArrayData_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("arrayData", ex.Message);
    }

    #endregion
}
