using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptPropertiesToolTests : TestBase
{
    private readonly PptPropertiesTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithProperties(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.DocumentProperties.Title = "Original Title";
        presentation.DocumentProperties.Author = "Original Author";
        presentation.DocumentProperties.Subject = "Original Subject";
        presentation.DocumentProperties.Keywords = "keyword1, keyword2";
        presentation.DocumentProperties.Category = "Test Category";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region Error Handling Tests

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
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Get Tests

    [Fact]
    public async Task Get_ShouldReturnPropertiesAsJson()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_properties.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("title", out _));
        Assert.True(json.RootElement.TryGetProperty("author", out _));
        Assert.True(json.RootElement.TryGetProperty("subject", out _));
        Assert.True(json.RootElement.TryGetProperty("createdTime", out _));
    }

    [Fact]
    public async Task Get_WithPresetProperties_ShouldReturnCorrectValues()
    {
        // Arrange
        var pptPath = CreatePresentationWithProperties("test_get_preset.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            Assert.True(json.RootElement.TryGetProperty("title", out _));
        }
        else
        {
            Assert.Equal("Original Title", json.RootElement.GetProperty("title").GetString());
            Assert.Equal("Original Author", json.RootElement.GetProperty("author").GetString());
            Assert.Equal("Original Subject", json.RootElement.GetProperty("subject").GetString());
        }
    }

    #endregion

    #region Set Tests

    [Fact]
    public async Task Set_ShouldSetBasicProperties()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_basic.pptx");
        var outputPath = CreateTestFilePath("test_set_basic_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["title"] = "Test Presentation",
            ["author"] = "Test Author",
            ["subject"] = "Test Subject"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Title", result);
        Assert.Contains("Author", result);
        Assert.Contains("Subject", result);

        using var presentation = new Presentation(outputPath);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            Assert.True(File.Exists(outputPath));
        }
        else
        {
            Assert.Equal("Test Presentation", presentation.DocumentProperties.Title);
            Assert.Equal("Test Author", presentation.DocumentProperties.Author);
            Assert.Equal("Test Subject", presentation.DocumentProperties.Subject);
        }
    }

    [Fact]
    public async Task Set_ShouldSetAllProperties()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_all.pptx");
        var outputPath = CreateTestFilePath("test_set_all_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["title"] = "Full Title",
            ["author"] = "Full Author",
            ["subject"] = "Full Subject",
            ["keywords"] = "key1, key2, key3",
            ["comments"] = "Test comments",
            ["category"] = "Test Category",
            ["company"] = "Test Company",
            ["manager"] = "Test Manager"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Title", result);
        Assert.Contains("Keywords", result);
        Assert.Contains("Company", result);
        Assert.Contains("Manager", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Set_WithCustomProperties_ShouldSetStringProperties()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_custom_string.pptx");
        var outputPath = CreateTestFilePath("test_set_custom_string_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["customProperties"] = new JsonObject
            {
                ["CustomKey1"] = "CustomValue1",
                ["CustomKey2"] = "CustomValue2"
            }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("CustomProperties", result);

        using var presentation = new Presentation(outputPath);
        var props = presentation.DocumentProperties;

        var isEvaluationMode = IsEvaluationMode();
        if (!isEvaluationMode)
        {
            Assert.Equal("CustomValue1", props["CustomKey1"]?.ToString());
            Assert.Equal("CustomValue2", props["CustomKey2"]?.ToString());
        }
    }

    [Fact]
    public async Task Set_WithCustomProperties_ShouldSetMultipleTypes()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_custom_types.pptx");
        var outputPath = CreateTestFilePath("test_set_custom_types_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["customProperties"] = new JsonObject
            {
                ["IntValue"] = 42,
                ["DoubleValue"] = 3.14,
                ["BoolValue"] = true,
                ["StringValue"] = "Hello"
            }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("CustomProperties", result);

        using var presentation = new Presentation(outputPath);
        var props = presentation.DocumentProperties;

        var isEvaluationMode = IsEvaluationMode();
        if (!isEvaluationMode)
        {
            Assert.Equal(42, props["IntValue"]);
            Assert.Equal(3.14, props["DoubleValue"]);
            Assert.Equal(true, props["BoolValue"]);
            Assert.Equal("Hello", props["StringValue"]);
        }
    }

    [Fact]
    public async Task Set_WithNoProperties_ShouldNotFail()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_empty.pptx");
        var outputPath = CreateTestFilePath("test_set_empty_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Output:", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion
}