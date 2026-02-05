using Aspose.Cells.Drawing;
using AsposeMcpServer.Handlers.Excel.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Shape;

public class EditExcelShapeHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelShapeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeEdit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_WithTextChange_ShouldUpdateText()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 0, 0, 0, 0, 100, 100);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "text", "Updated Text" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Updated Text", worksheet.Shapes[0].Text);
    }

    [Fact]
    public void Execute_WithSizeChange_ShouldUpdateSize()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 0, 0, 0, 0, 100, 100);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "width", 200 },
            { "height", 150 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(200, worksheet.Shapes[0].Width);
        Assert.Equal(150, worksheet.Shapes[0].Height);
    }

    [Fact]
    public void Execute_WithNoChanges_ShouldReturnNoChangesMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 0, 0, 0, 0, 100, 100);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("No changes", result.Message);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 0, 0, 0, 0, 100, 100);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "text", "Modified" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingShapeIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

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
}
