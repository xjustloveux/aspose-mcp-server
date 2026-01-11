using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

public class ApplyThemeHandlerTests : PptHandlerTestBase
{
    private readonly ApplyThemeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ApplyTheme()
    {
        Assert.Equal("apply_theme", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutThemePath_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentThemePath_ThrowsFileNotFoundException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "themePath", @"C:\nonexistent\theme.pptx" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
