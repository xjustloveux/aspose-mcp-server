using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptSmartArtToolTests : TestBase
{
    private readonly PptSmartArtTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task AddSmartArt_ShouldAddSmartArt()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_smartart.pptx");
        var outputPath = CreateTestFilePath("test_add_smartart_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["layout"] = "BasicProcess",
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 400,
            ["height"] = 300
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var smartArts = slide.Shapes.OfType<ISmartArt>().ToList();
        Assert.True(smartArts.Count > 0, "Slide should contain at least one SmartArt");
    }

    [Fact]
    public async Task GetSmartArt_ShouldReturnSmartArtInfo()
    {
        // Arrange - PptSmartArtTool doesn't have a "get" operation, test manage_nodes instead
        var pptPath = CreateTestPresentation("test_manage_smartart_nodes.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var smartArt = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            presentation.Save(pptPath, SaveFormat.Pptx);

            // Find the index of the SmartArt shape (it might not be at index 0 if there are placeholders)
            smartArtShapeIndex = -1;
            for (var i = 0; i < slide.Shapes.Count; i++)
                if (slide.Shapes[i] == smartArt)
                {
                    smartArtShapeIndex = i;
                    break;
                }

            Assert.True(smartArtShapeIndex >= 0, "SmartArt shape should be found in slide");
        }

        var outputPath = CreateTestFilePath("test_manage_smartart_nodes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = smartArtShapeIndex,
            ["action"] = "add",
            ["targetPath"] = new JsonArray { 0 },
            ["text"] = "New Node"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("SmartArt", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AddSmartArt_InvalidLayout_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_invalid_layout.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["layout"] = "InvalidLayoutName"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ManageNodes_InvalidRootIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_invalid_root_index.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = smartArtShapeIndex,
            ["action"] = "edit",
            ["targetPath"] = new JsonArray { 999 },
            ["text"] = "Test"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ManageNodes_EditNode_ShouldEditText()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_edit_node.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_node_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = smartArtShapeIndex,
            ["action"] = "edit",
            ["targetPath"] = new JsonArray { 0 },
            ["text"] = "Edited Text"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Edited Text", result);
    }

    [Fact]
    public async Task ManageNodes_DeleteRootNode_ShouldSucceed()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_delete_root_node.pptx");
        int smartArtShapeIndex;
        int initialNodeCount;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var smartArt = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            initialNodeCount = smartArt.AllNodes.Count;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_root_node_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = smartArtShapeIndex,
            ["action"] = "delete",
            ["targetPath"] = new JsonArray { 0 }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        using var resultPres = new Presentation(outputPath);
        var resultSmartArt = resultPres.Slides[0].Shapes[smartArtShapeIndex] as ISmartArt;
        Assert.NotNull(resultSmartArt);
        Assert.True(resultSmartArt.AllNodes.Count < initialNodeCount);
    }

    [Fact]
    public async Task UnknownOperation_ShouldThrow()
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

    [Fact]
    public async Task ManageNodes_NotSmartArtShape_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_not_smartart.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0,
            ["action"] = "edit",
            ["targetPath"] = new JsonArray { 0 },
            ["text"] = "Test"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ManageNodes_AddWithPosition_ShouldInsertAtPosition()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_add_with_position.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_add_with_position_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = smartArtShapeIndex,
            ["action"] = "add",
            ["targetPath"] = new JsonArray { 0 },
            ["text"] = "Inserted Node",
            ["position"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("position 0", result);
        Assert.Contains("Inserted Node", result);
    }

    [Fact]
    public async Task ManageNodes_AddWithInvalidPosition_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_add_invalid_position.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = smartArtShapeIndex,
            ["action"] = "add",
            ["targetPath"] = new JsonArray { 0 },
            ["text"] = "Test",
            ["position"] = 999
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}