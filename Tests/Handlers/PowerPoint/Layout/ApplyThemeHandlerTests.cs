using Aspose.Slides;
using Aspose.Slides.Export;
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

    #region Theme Application

    [SkippableFact]
    public void Execute_WithValidTheme_AppliesTheme()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode limits master slide operations");

        var themePath = Path.Combine(Path.GetTempPath(), $"theme_{Guid.NewGuid()}.pptx");
        try
        {
            using (var themePresentation = new Presentation())
            {
                themePresentation.Save(themePath, SaveFormat.Pptx);
            }

            var pres = CreateEmptyPresentation();
            var context = CreateContext(pres);
            var parameters = CreateParameters(new Dictionary<string, object?>
            {
                { "themePath", themePath }
            });

            var result = _handler.Execute(context, parameters);

            Assert.Contains("Theme applied", result);
            Assert.Contains("master", result.ToLower());
            AssertModified(context);
        }
        finally
        {
            if (File.Exists(themePath)) File.Delete(themePath);
        }
    }

    [SkippableFact]
    public void Execute_WithSlides_AppliesLayoutToAllSlides()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode limits slide operations");

        var themePath = Path.Combine(Path.GetTempPath(), $"theme_slides_{Guid.NewGuid()}.pptx");
        try
        {
            using (var themePresentation = new Presentation())
            {
                themePresentation.Save(themePath, SaveFormat.Pptx);
            }

            var pres = CreateEmptyPresentation();
            pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);
            pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);
            var context = CreateContext(pres);
            var parameters = CreateParameters(new Dictionary<string, object?>
            {
                { "themePath", themePath }
            });

            var result = _handler.Execute(context, parameters);

            Assert.Contains("Theme applied", result);
            Assert.Contains("layout applied to all slides", result);
        }
        finally
        {
            if (File.Exists(themePath)) File.Delete(themePath);
        }
    }

    #endregion
}
