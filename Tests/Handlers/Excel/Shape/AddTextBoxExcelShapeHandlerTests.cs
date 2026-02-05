using AsposeMcpServer.Handlers.Excel.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Shape;

public class AddTextBoxExcelShapeHandlerTests : ExcelHandlerTestBase
{
    private readonly AddTextBoxExcelShapeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeAddTextbox()
    {
        Assert.Equal("add_textbox", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingText_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_ShouldAddTextBox()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test TextBox" }
        });

        _handler.Execute(context, parameters);

        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Shapes);
        Assert.Equal("Test TextBox", worksheet.Shapes[0].Text);
    }

    [Fact]
    public void Execute_WithCustomSize_ShouldSetSize()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Sized TextBox" },
            { "width", 300 },
            { "height", 100 }
        });

        _handler.Execute(context, parameters);

        var worksheet = workbook.Worksheets[0];
        Assert.Equal(300, worksheet.Shapes[0].Width);
        Assert.Equal(100, worksheet.Shapes[0].Height);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Modified TextBox" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion
}
