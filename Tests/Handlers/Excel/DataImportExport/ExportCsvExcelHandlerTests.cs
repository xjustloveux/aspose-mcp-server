using AsposeMcpServer.Handlers.Excel.DataImportExport;
using AsposeMcpServer.Results.Excel.DataImportExport;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataImportExport;

public class ExportCsvExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly ExportCsvExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeExportCsv()
    {
        Assert.Equal("export_csv", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithValidData_ShouldExportCsv()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Age" },
            { "John", 30 },
            { "Jane", 25 }
        });
        var outputPath = CreateTempFile(".csv", "");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExportExcelResult>(res);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(result.RowCount > 0);
        Assert.Contains("exported to CSV", result.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_ShouldNotMarkModified()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Age" },
            { "John", 30 }
        });
        var outputPath = CreateTempFile(".csv", "");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithMissingOutputPath_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("outputPath", ex.Message);
    }

    #endregion
}
