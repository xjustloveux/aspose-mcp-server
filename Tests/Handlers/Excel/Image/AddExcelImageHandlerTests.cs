using AsposeMcpServer.Handlers.Excel.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Image;

public class AddExcelImageHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_AddsToCorrectSheet()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "imagePath", tempFile },
            { "cell", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(workbook.Worksheets[0].Pictures);
        Assert.Single(workbook.Worksheets[1].Pictures);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsImage()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Image added", result.Message);
        Assert.Single(workbook.Worksheets[0].Pictures);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsCellAddress()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "cell", "B5" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B5", result.Message);
    }

    [Fact]
    public void Execute_ReturnsImageSize()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("size:", result.Message);
    }

    #endregion

    #region Image Sizing Options

    [Fact]
    public void Execute_WithWidthAndHeight_SetsSize()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "cell", "A1" },
            { "width", 200 },
            { "height", 150 }
        });

        _handler.Execute(context, parameters);

        var picture = workbook.Worksheets[0].Pictures[0];
        Assert.Equal(200, picture.Width);
        Assert.Equal(150, picture.Height);
    }

    [Fact]
    public void Execute_WithKeepAspectRatioFalse_SetsFlag()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "cell", "A1" },
            { "width", 100 },
            { "keepAspectRatio", false }
        });

        _handler.Execute(context, parameters);

        var picture = workbook.Worksheets[0].Pictures[0];
        Assert.False(picture.IsLockAspectRatio);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutImagePath_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imagePath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "/nonexistent/path/image.png" },
            { "cell", "A1" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var tempFile = CreateTempImageFile();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "imagePath", tempFile },
            { "cell", "A1" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
