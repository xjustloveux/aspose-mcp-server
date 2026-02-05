using Aspose.Cells.Drawing;
using AsposeMcpServer.Handlers.Excel.Shape;
using AsposeMcpServer.Results.Excel.Shape;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Shape;

public class GetExcelShapesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelShapesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeGet()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidShapeIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoShapes_ShouldReturnEmpty()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesExcelResult>(res);
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Execute_WithShape_ShouldReturnInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 0, 0, 0, 0, 100, 100);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesExcelResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Single(result.Items);
    }

    [Fact]
    public void Execute_WithSpecificShapeIndex_ShouldReturnSingle()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 0, 0, 0, 0, 100, 100);
        worksheet.Shapes.AddAutoShape(AutoShapeType.Oval, 5, 0, 5, 0, 80, 80);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesExcelResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Single(result.Items);
        Assert.Equal(1, result.Items[0].Index);
    }

    #endregion
}
