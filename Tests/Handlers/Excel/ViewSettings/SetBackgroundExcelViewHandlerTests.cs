using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetBackgroundExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetBackgroundExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetBackground()
    {
        Assert.Equal("set_background", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_RemovesBackground()
    {
        var workbook = CreateWorkbookWithBackground();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "removeBackground", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        Assert.Null(workbook.Worksheets[0].BackgroundImage);
        AssertModified(context);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithBackground()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].BackgroundImage = [0x89, 0x50, 0x4E, 0x47];
        return workbook;
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNoParameters_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentImagePath_ThrowsFileNotFoundException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "nonexistent.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
