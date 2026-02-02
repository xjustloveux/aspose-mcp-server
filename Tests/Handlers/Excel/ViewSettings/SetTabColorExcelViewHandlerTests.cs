using System.Drawing;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetTabColorExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetTabColorExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetTabColor()
    {
        Assert.Equal("set_tab_color", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutColor_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsTabColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "Red" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var tabColor = workbook.Worksheets[0].TabColor;
            Assert.Equal(Color.Red.R, tabColor.R);
            Assert.Equal(Color.Red.G, tabColor.G);
            Assert.Equal(Color.Red.B, tabColor.B);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsTabColorOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "color", "Blue" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var tabColor = workbook.Worksheets[1].TabColor;
            Assert.Equal(Color.Blue.R, tabColor.R);
            Assert.Equal(Color.Blue.G, tabColor.G);
            Assert.Equal(Color.Blue.B, tabColor.B);
        }
    }

    [Fact]
    public void Execute_WithHexColor_SetsTabColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "#FF5733" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var tabColor = workbook.Worksheets[0].TabColor;
            Assert.Equal(0xFF, tabColor.R);
            Assert.Equal(0x57, tabColor.G);
            Assert.Equal(0x33, tabColor.B);
        }

        AssertModified(context);
    }

    #endregion
}
