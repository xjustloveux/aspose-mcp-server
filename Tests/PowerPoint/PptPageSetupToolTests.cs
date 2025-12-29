using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptPageSetupToolTests : TestBase
{
    private readonly PptPageSetupTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetSlideSize_WithPreset_ShouldSetSlideSize()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_slide_size.pptx");
        var outputPath = CreateTestFilePath("test_set_slide_size_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_size",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["preset"] = "OnScreen16x9"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Slide size set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideSizeType.OnScreen, presentation.SlideSize.Type);
    }

    [Fact]
    public async Task SetSlideSize_WithCustomSize_ShouldSetCustomDimensions()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_custom_size.pptx");
        var outputPath = CreateTestFilePath("test_set_custom_size_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_size",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["preset"] = "Custom",
            ["width"] = 720,
            ["height"] = 540
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Custom", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(720, presentation.SlideSize.Size.Width, 1);
        Assert.Equal(540, presentation.SlideSize.Size.Height, 1);
    }

    [Fact]
    public async Task SetSlideSize_CustomWithoutDimensions_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_custom_no_dims.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_size",
            ["path"] = pptPath,
            ["preset"] = "Custom"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SetSlideSize_WithOutOfRangeWidth_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_size_out_of_range.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_size",
            ["path"] = pptPath,
            ["preset"] = "Custom",
            ["width"] = 10000,
            ["height"] = 540
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SetSlideOrientation_Portrait_ShouldSwapWidthAndHeight()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_orientation.pptx");
        var outputPath = CreateTestFilePath("test_set_orientation_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "set_orientation",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["orientation"] = "Portrait"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Portrait", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.SlideSize.Size.Height > presentation.SlideSize.Size.Width);
    }

    [Fact]
    public async Task SetSlideOrientation_Landscape_ShouldKeepOrSwapToLandscape()
    {
        // Arrange - Create a portrait presentation first
        var pptPath = CreateTestFilePath("test_landscape_orientation.pptx");
        using (var ppt = new Presentation())
        {
            ppt.SlideSize.SetSize(540, 720, SlideSizeScaleType.DoNotScale);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_landscape_orientation_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_orientation",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["orientation"] = "Landscape"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Landscape", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.SlideSize.Size.Width > presentation.SlideSize.Size.Height);
    }

    [Fact]
    public async Task SetFooter_ShouldSetFooterText()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_footer.pptx");
        var outputPath = CreateTestFilePath("test_set_footer_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_footer",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["footerText"] = "Footer Text"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetFooter_WithDateText_ShouldSetDateText()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_footer_date.pptx");
        var outputPath = CreateTestFilePath("test_footer_date_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_footer",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["footerText"] = "Footer",
            ["dateText"] = "2024-12-28"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetFooter_WithSlideIndices_ShouldSetForSpecificSlides()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_footer_indices.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_footer_indices_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_footer",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["footerText"] = "Footer for specific slides",
            ["slideIndices"] = new JsonArray { 0, 1 }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("2 slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetFooter_InvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_footer_invalid.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_footer",
            ["path"] = pptPath,
            ["footerText"] = "Test",
            ["slideIndices"] = new JsonArray { 99 }
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SetSlideNumbering_ShouldSetSlideNumbering()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_slide_numbering.pptx");
        var outputPath = CreateTestFilePath("test_slide_numbering_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_slide_numbering",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["showSlideNumber"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public async Task SetSlideNumbering_WithFirstNumber_ShouldSetStartingNumber()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_slide_numbering_first.pptx");
        var outputPath = CreateTestFilePath("test_slide_numbering_first_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_slide_numbering",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["showSlideNumber"] = true,
            ["firstNumber"] = 5
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("starting from 5", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(5, presentation.FirstSlideNumber);
    }

    [Fact]
    public async Task SetSlideNumbering_Hide_ShouldHideNumbers()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_hide_numbering.pptx");
        var outputPath = CreateTestFilePath("test_hide_numbering_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_slide_numbering",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["showSlideNumber"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("hidden", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}