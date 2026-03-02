using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

[SupportedOSPlatform("windows")]
public class ApplyThemeHandlerTests : PptHandlerTestBase
{
    private readonly ApplyThemeHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_ApplyTheme()
    {
        SkipIfNotWindows();
        Assert.Equal("apply_theme", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutThemePath_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithNonExistentThemePath_ThrowsFileNotFoundException()
    {
        SkipIfNotWindows();
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
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode limits master slide operations");

        var themePath = Path.Combine(Path.GetTempPath(), $"theme_{Guid.NewGuid()}.pptx");
        try
        {
            using (var themePresentation = new Presentation())
            {
                themePresentation.Save(themePath, SaveFormat.Pptx);
            }

            var pres = CreateEmptyPresentation();
            var originalMasterCount = pres.Masters.Count;
            var context = CreateContext(pres);
            var parameters = CreateParameters(new Dictionary<string, object?>
            {
                { "themePath", themePath }
            });

            var res = _handler.Execute(context, parameters);

            Assert.IsType<SuccessResult>(res);
            Assert.True(pres.Masters.Count > originalMasterCount,
                "Master slide count should increase after applying theme");
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
        SkipIfNotWindows();
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
            var originalMasterCount = pres.Masters.Count;
            var context = CreateContext(pres);
            var parameters = CreateParameters(new Dictionary<string, object?>
            {
                { "themePath", themePath }
            });

            var res = _handler.Execute(context, parameters);

            Assert.IsType<SuccessResult>(res);
            Assert.True(pres.Masters.Count > originalMasterCount,
                "Master slide count should increase after applying theme");
            var newMaster = pres.Masters[^1];
            var expectedLayout = newMaster.LayoutSlides[0];
            foreach (var slide in pres.Slides) Assert.Equal(expectedLayout, slide.LayoutSlide);
        }
        finally
        {
            if (File.Exists(themePath)) File.Delete(themePath);
        }
    }

    #endregion
}
