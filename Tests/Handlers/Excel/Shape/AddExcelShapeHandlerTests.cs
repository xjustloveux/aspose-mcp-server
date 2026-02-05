using Aspose.Cells.Drawing;
using AsposeMcpServer.Handlers.Excel.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Shape;

public class AddExcelShapeHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelShapeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeAdd()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region ResolveAutoShapeType

    [Theory]
    [InlineData("Rectangle", AutoShapeType.Rectangle)]
    [InlineData("Oval", AutoShapeType.Oval)]
    [InlineData("Star5", AutoShapeType.Star5)]
    [InlineData("Diamond", AutoShapeType.Diamond)]
    public void ResolveAutoShapeType_WithValidTypes_ShouldReturn(string shapeType, AutoShapeType expected)
    {
        var result = AddExcelShapeHandler.ResolveAutoShapeType(shapeType);

        Assert.Equal(expected, result);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_WithRectangle_ShouldAddShape()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Rectangle" }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].Shapes);
    }

    [Fact]
    public void Execute_WithText_ShouldSetText()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Rectangle" },
            { "text", "Hello Shape" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Hello Shape", workbook.Worksheets[0].Shapes[0].Text);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Rectangle" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingShapeType_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeType", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidShapeType_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "NonExistentShape" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
