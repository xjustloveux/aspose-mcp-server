using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.FileOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.FileOperations;

public class CreateWorkbookHandlerTests : ExcelHandlerTestBase
{
    private readonly CreateWorkbookHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPathOrOutputPath_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_WithPath_CreatesWorkbook()
    {
        var outputPath = Path.Combine(TestDir, "test.xlsx");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created successfully", result.ToLower());
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created workbook should have content");
    }

    [Fact]
    public void Execute_WithOutputPath_CreatesWorkbook()
    {
        var outputPath = Path.Combine(TestDir, "output.xlsx");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created successfully", result.ToLower());
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created workbook should have content");
    }

    [Fact]
    public void Execute_WithSheetName_CreatesWorkbookWithNamedSheet()
    {
        var outputPath = Path.Combine(TestDir, "named_sheet.xlsx");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath },
            { "sheetName", "MySheet" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created successfully", result.ToLower());
        Assert.True(File.Exists(outputPath));

        using var createdWorkbook = new Workbook(outputPath);
        Assert.Equal("MySheet", createdWorkbook.Worksheets[0].Name);
    }

    #endregion
}
