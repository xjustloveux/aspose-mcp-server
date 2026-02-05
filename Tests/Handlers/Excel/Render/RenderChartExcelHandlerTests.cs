using AsposeMcpServer.Handlers.Excel.Render;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Render;

public class RenderChartExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly RenderChartExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeRenderChart()
    {
        Assert.Equal("render_chart", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithMissingOutputPath_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("outputPath", ex.Message);
    }

    [Fact]
    public void Execute_WithMissingChartIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.png");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("chartIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidChartIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.png");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "chartIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
