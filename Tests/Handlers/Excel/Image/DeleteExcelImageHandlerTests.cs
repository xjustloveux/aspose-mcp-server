using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Image;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Image;

public class DeleteExcelImageHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutImageIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(0)]
    [InlineData(10)]
    public void Execute_WithInvalidImageIndex_ThrowsArgumentException(int invalidIndex)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", invalidIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "imageIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Delete With Actual Images

    [Fact]
    public void Execute_DeletesImage()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateWorkbookWithImage(tempFile);
        Assert.Single(workbook.Worksheets[0].Pictures);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result);
        Assert.Empty(workbook.Worksheets[0].Pictures);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateWorkbookWithMultipleImages(tempFile, 3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2 images remaining", result);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithImage(string imagePath)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Pictures.Add(0, 0, imagePath);
        return workbook;
    }

    private static Workbook CreateWorkbookWithMultipleImages(string imagePath, int count)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        for (var i = 0; i < count; i++) sheet.Pictures.Add(i, 0, imagePath);
        return workbook;
    }

    #endregion
}
