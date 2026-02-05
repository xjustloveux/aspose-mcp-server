using AsposeMcpServer.Handlers.Excel.DataImportExport;
using AsposeMcpServer.Results.Excel.DataImportExport;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataImportExport;

public class ImportJsonExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly ImportJsonExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeImportJson()
    {
        Assert.Equal("import_json", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithJsonArray_ShouldImportData()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "jsonData", "[{\"name\":\"John\",\"age\":30},{\"name\":\"Jane\",\"age\":25}]" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ImportExcelResult>(res);
        Assert.True(result.RowCount > 0);
        Assert.True(result.ColumnCount > 0);
        Assert.Equal("A1", result.StartCell);
        Assert.Contains("JSON data imported", result.Message);
    }

    [Fact]
    public void Execute_WithStartCell_ShouldImportAtPosition()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "jsonData", "[{\"name\":\"John\",\"age\":30}]" },
            { "startCell", "C3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ImportExcelResult>(res);
        Assert.Equal("C3", result.StartCell);
        Assert.Contains("C3", result.Message);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "jsonData", "[{\"name\":\"John\",\"age\":30}]" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMissingJsonData_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("jsonData", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyJsonData_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "jsonData", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("jsonData", ex.Message);
    }

    #endregion
}
