using Aspose.Slides;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Handlers.PowerPoint.SmartArt;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.SmartArt;

public class ManageSmartArtNodesHandlerTests : PptHandlerTestBase
{
    private readonly ManageSmartArtNodesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ManageNodes()
    {
        Assert.Equal("manage_nodes", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithSmartArt()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
        return pres;
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_AddNode_AddsNewNode()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "New Node" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("node added", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditNode_EditsExistingNode()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "edit" },
            { "targetPath", "[0]" },
            { "text", "Updated Text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeleteNode_DeletesNode()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "delete" },
            { "targetPath", "[0]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutAction_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "targetPath", "[0]" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidAction_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "invalid" },
            { "targetPath", "[0]" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonSmartArtShape_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Test");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
