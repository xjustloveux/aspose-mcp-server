using AsposeMcpServer.Handlers.Excel.Image;
using AsposeMcpServer.Results.Excel.Image;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Image;

public class GetExcelImagesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelImagesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_ReturnsCorrectSheetInfo()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("ImageSheet");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesExcelResult>(res);

        Assert.Equal("ImageSheet", result.WorksheetName);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoImages_ReturnsEmptyList()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesExcelResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Contains("No images found", result.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesExcelResult>(res);

        Assert.Equal("Sheet1", result.WorksheetName);
    }

    [Fact]
    public void Execute_ReturnsValidJsonStructure()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesExcelResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.WorksheetName);
        Assert.NotNull(result.Items);
    }

    #endregion
}
