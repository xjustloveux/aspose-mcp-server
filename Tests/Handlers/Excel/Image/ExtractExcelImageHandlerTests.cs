using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Image;

public class ExtractExcelImageHandlerTests : ExcelHandlerTestBase
{
    private readonly ExtractExcelImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Extract()
    {
        Assert.Equal("extract", _handler.Operation);
    }

    #endregion

    #region Export Formats

    [Theory]
    [InlineData(".png")]
    [InlineData(".jpg")]
    [InlineData(".bmp")]
    public void Execute_WithVariousFormats_ExtractsSuccessfully(string extension)
    {
        var tempImageFile = CreateTempImageFile();
        var exportPath = Path.Combine(TestDir, $"extracted_{Guid.NewGuid()}{extension}");
        var workbook = CreateWorkbookWithImage(tempImageFile);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "exportPath", exportPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("extracted", result.Message);
        Assert.True(File.Exists(exportPath));
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

    #endregion

    #region Basic Extract Operations

    [Fact]
    public void Execute_ExtractsImage()
    {
        var tempImageFile = CreateTempImageFile();
        var exportPath = Path.Combine(TestDir, $"extracted_{Guid.NewGuid()}.png");
        var workbook = CreateWorkbookWithImage(tempImageFile);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "exportPath", exportPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("extracted", result.Message);
        Assert.True(File.Exists(exportPath));
        var fileInfo = new FileInfo(exportPath);
        Assert.True(fileInfo.Length > 0, "Extracted image should have content");
    }

    [Fact]
    public void Execute_ReturnsImageIndex()
    {
        var tempImageFile = CreateTempImageFile();
        var exportPath = Path.Combine(TestDir, $"extracted_{Guid.NewGuid()}.png");
        var workbook = CreateWorkbookWithImage(tempImageFile);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "exportPath", exportPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#0", result.Message);
    }

    [Fact]
    public void Execute_ReturnsExportPath()
    {
        var tempImageFile = CreateTempImageFile();
        var exportPath = Path.Combine(TestDir, $"extracted_{Guid.NewGuid()}.png");
        var workbook = CreateWorkbookWithImage(tempImageFile);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "exportPath", exportPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(exportPath, result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutImageIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "exportPath", "/tmp/test.png" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutExportPath_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("exportPath", ex.Message, StringComparison.OrdinalIgnoreCase);
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
            { "imageIndex", invalidIndex },
            { "exportPath", "/tmp/test.png" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var tempImageFile = CreateTempImageFile();
        var workbook = CreateWorkbookWithImage(tempImageFile);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "exportPath", "/tmp/test.xyz" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported export format", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "imageIndex", 0 },
            { "exportPath", "/tmp/test.png" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
