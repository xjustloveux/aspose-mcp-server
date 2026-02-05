using Aspose.Cells.Drawing;
using AsposeMcpServer.Handlers.Excel.Render;
using AsposeMcpServer.Results.Excel.Render;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Render;

public class RenderSheetExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly RenderSheetExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeRenderSheet()
    {
        Assert.Equal("render_sheet", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithValidData_ShouldRenderSheet()
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

            var result = Assert.IsType<RenderExcelResult>(res);
            Assert.NotEmpty(result.OutputPaths);
            Assert.True(result.PageCount >= 0);
            Assert.Equal("png", result.Format);
            Assert.Contains("rendered", result.Message);
        }
        finally
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);
        }
    }

    [Fact]
    public void Execute_WithJpegFormat_ShouldRender()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Age" },
            { "John", 30 }
        });
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.jpg");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "format", "jpeg" }
        });

        try
        {
            var res = _handler.Execute(context, parameters);

            var result = Assert.IsType<RenderExcelResult>(res);
            Assert.Equal("jpeg", result.Format);
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
        var result = RenderSheetExcelHandler.ResolveImageType(format);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ResolveImageType_WithInvalidFormat_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            RenderSheetExcelHandler.ResolveImageType("invalid"));
        Assert.Contains("Unknown image format", ex.Message);
    }

    #endregion
}
