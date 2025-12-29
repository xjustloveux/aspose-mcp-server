using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptTransitionToolTests : TestBase
{
    private readonly PptTransitionTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetTransition_ShouldSetTransition()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_transition.pptx");
        var outputPath = CreateTestFilePath("test_set_transition_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["transitionType"] = "Fade",
            ["advanceAfterSeconds"] = 1.5
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Fade, slide.SlideShowTransition.Type);
        Assert.Equal(1500u, slide.SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public async Task GetTransition_ShouldReturnTransition()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_transition.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            slide.SlideShowTransition.AdvanceAfterTime = 2000;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("\"type\": \"Fade\"", result);
        Assert.Contains("\"hasTransition\": true", result);
        Assert.Contains("\"advanceAfterSeconds\": 2", result);
    }

    [Fact]
    public async Task DeleteTransition_ShouldDeleteTransition()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_transition.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_transition_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        Assert.Equal(TransitionType.None, resultSlide.SlideShowTransition.Type);
        Assert.Equal(0u, resultSlide.SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public async Task SetTransition_WithEnumTryParse_ShouldSupportAllTypes()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_transition_push.pptx");
        var outputPath = CreateTestFilePath("test_set_transition_push_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["transitionType"] = "Push"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Push, slide.SlideShowTransition.Type);
    }

    [Fact]
    public async Task SetTransition_WithInvalidType_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_transition_invalid.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["transitionType"] = "InvalidType"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SetTransition_WithoutAdvanceAfterSeconds_ShouldNotSetAutoAdvance()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_transition_no_advance.pptx");
        var outputPath = CreateTestFilePath("test_set_transition_no_advance_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["transitionType"] = "Wipe"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Wipe, slide.SlideShowTransition.Type);
    }

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}