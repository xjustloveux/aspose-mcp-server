using Aspose.Cells.Drawing;
using AsposeMcpServer.Handlers.Excel.DataImportExport;
using AsposeMcpServer.Results.Excel.DataImportExport;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataImportExport;

public class ExportRangeImageExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly ExportRangeImageExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeExportRangeImage()
    {
        Assert.Equal("export_range_image", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithValidData_ShouldExportImage()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Age" },
            { "John", 30 },
            { "Jane", 25 }
        });
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.png");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        try
        {
            var res = _handler.Execute(context, parameters);

            var result = Assert.IsType<ExportExcelResult>(res);
            Assert.Equal(outputPath, result.OutputPath);
            Assert.Contains("exported to image", result.Message);
        }
        finally
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);
        }
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

    #region ResolveImageType Tests

    [Theory]
    [InlineData("png", ImageType.Png)]
    [InlineData("jpeg", ImageType.Jpeg)]
    [InlineData("bmp", ImageType.Bmp)]
    [InlineData("tiff", ImageType.Tiff)]
    [InlineData("svg", ImageType.Svg)]
    public void ResolveImageType_WithValidFormats_ShouldReturn(string format, ImageType expected)
    {
        var result = ExportRangeImageExcelHandler.ResolveImageType(format);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ResolveImageType_WithInvalidFormat_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExportRangeImageExcelHandler.ResolveImageType("invalid"));
        Assert.Contains("Unknown image format", ex.Message);
    }

    #endregion
}
