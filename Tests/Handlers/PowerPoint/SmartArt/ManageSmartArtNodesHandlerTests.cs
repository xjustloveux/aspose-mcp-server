using Aspose.Slides;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Handlers.PowerPoint.SmartArt;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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
        var smartArt = (ISmartArt)pres.Slides[0].Shapes[0];
        var targetNode = smartArt.AllNodes[0];
        var initialChildCount = targetNode.ChildNodes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "New Node" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.Equal(initialChildCount + 1, targetNode.ChildNodes.Count);
            var newNode = targetNode.ChildNodes[^1];
            Assert.Equal("New Node", newNode.TextFrame.Text);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var smartArt = (ISmartArt)pres.Slides[0].Shapes[0];
            Assert.Equal("Updated Text", smartArt.AllNodes[0].TextFrame.Text);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_DeleteNode_DeletesNode()
    {
        var pres = CreatePresentationWithSmartArt();
        var smartArt = (ISmartArt)pres.Slides[0].Shapes[0];
        var initialNodeCount = smartArt.AllNodes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "delete" },
            { "targetPath", "[0]" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
            Assert.True(smartArt.AllNodes.Count < initialNodeCount);
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

    [Fact]
    public void Execute_WithoutTargetPath_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidRootIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "edit" },
            { "targetPath", "[99]" },
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_AddWithoutText_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_EditWithoutText_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "edit" },
            { "targetPath", "[0]" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    #endregion

    #region Position Parameter Tests

    [Fact]
    public void Execute_AddWithPosition_AddsAtPosition()
    {
        var pres = CreatePresentationWithSmartArt();
        var smartArt = (ISmartArt)pres.Slides[0].Shapes[0];
        var targetNode = smartArt.AllNodes[0];
        var initialChildCount = targetNode.ChildNodes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "Positioned Node" },
            { "position", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.Equal(initialChildCount + 1, targetNode.ChildNodes.Count);
            Assert.Equal("Positioned Node", targetNode.ChildNodes[0].TextFrame.Text);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_AddWithInvalidPosition_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "Test" },
            { "position", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Position", ex.Message);
    }

    #endregion

    #region Invalid Index Tests

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 99 },
            { "action", "add" },
            { "targetPath", "[0]" },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyTargetPath_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSmartArt();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "action", "add" },
            { "targetPath", "[]" },
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("at least one index", ex.Message);
    }

    #endregion
}
