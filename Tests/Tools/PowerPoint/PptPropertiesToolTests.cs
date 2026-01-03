using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptPropertiesToolTests : TestBase
{
    private readonly PptPropertiesTool _tool;

    public PptPropertiesToolTests()
    {
        _tool = new PptPropertiesTool(SessionManager);
    }

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

    #region General Tests

    [Fact]
    public void Get_ShouldReturnPropertiesAsJson()
    {
        var pptPath = CreateTestPresentation("test_get_properties.pptx");
        var result = _tool.Execute("get", pptPath);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("title", out _));
        Assert.True(json.RootElement.TryGetProperty("author", out _));
        Assert.True(json.RootElement.TryGetProperty("subject", out _));
        Assert.True(json.RootElement.TryGetProperty("createdTime", out _));
    }

    [Fact]
    public void Get_WithPresetProperties_ShouldReturnCorrectValues()
    {
        var pptPath = CreatePresentationWithProperties("test_get_preset.pptx");
        var result = _tool.Execute("get", pptPath);
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

    [Fact]
    public void Set_ShouldSetBasicProperties()
    {
        var pptPath = CreateTestPresentation("test_set_basic.pptx");
        var outputPath = CreateTestFilePath("test_set_basic_output.pptx");
        var result = _tool.Execute("set", pptPath, title: "Test Presentation", author: "Test Author",
            subject: "Test Subject", outputPath: outputPath);
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
    public void Set_ShouldSetAllProperties()
    {
        var pptPath = CreateTestPresentation("test_set_all.pptx");
        var outputPath = CreateTestFilePath("test_set_all_output.pptx");
        var result = _tool.Execute("set", pptPath,
            title: "Full Title",
            author: "Full Author",
            subject: "Full Subject",
            keywords: "key1, key2, key3",
            comments: "Test comments",
            category: "Test Category",
            company: "Test Company",
            manager: "Test Manager",
            outputPath: outputPath);
        Assert.Contains("Title", result);
        Assert.Contains("Keywords", result);
        Assert.Contains("Company", result);
        Assert.Contains("Manager", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Set_WithCustomProperties_ShouldSetStringProperties()
    {
        var pptPath = CreateTestPresentation("test_set_custom_string.pptx");
        var outputPath = CreateTestFilePath("test_set_custom_string_output.pptx");
        var customProperties = new Dictionary<string, object>
        {
            ["CustomKey1"] = "CustomValue1",
            ["CustomKey2"] = "CustomValue2"
        };
        var result = _tool.Execute("set", pptPath, customProperties: customProperties, outputPath: outputPath);
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
    public void Set_WithCustomProperties_ShouldSetMultipleTypes()
    {
        var pptPath = CreateTestPresentation("test_set_custom_types.pptx");
        var outputPath = CreateTestFilePath("test_set_custom_types_output.pptx");
        var customProperties = new Dictionary<string, object>
        {
            ["IntValue"] = 42,
            ["DoubleValue"] = 3.14,
            ["BoolValue"] = true,
            ["StringValue"] = "Hello"
        };
        var result = _tool.Execute("set", pptPath, customProperties: customProperties, outputPath: outputPath);
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
    public void Set_WithNoProperties_ShouldNotFail()
    {
        var pptPath = CreateTestPresentation("test_set_empty.pptx");
        var outputPath = CreateTestFilePath("test_set_empty_output.pptx");
        var result = _tool.Execute("set", pptPath, outputPath: outputPath);
        Assert.Contains("Output:", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Get_MissingPath_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get"));
        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldReturnPropertiesFromMemory()
    {
        var pptPath = CreatePresentationWithProperties("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("title", out _));
        Assert.True(json.RootElement.TryGetProperty("author", out _));
    }

    [Fact]
    public void Set_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, title: "Session Title", author: "Session Author");
        Assert.Contains("Title", result);
        Assert.Contains("session", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var isEvaluationMode = IsEvaluationMode();
        if (!isEvaluationMode)
        {
            Assert.Equal("Session Title", ppt.DocumentProperties.Title);
            Assert.Equal("Session Author", ppt.DocumentProperties.Author);
        }
    }

    [Fact]
    public void Set_WithSessionId_CustomProperties_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_custom.pptx");
        var sessionId = OpenSession(pptPath);
        var customProperties = new Dictionary<string, object>
        {
            ["SessionKey"] = "SessionValue"
        };
        var result = _tool.Execute("set", sessionId: sessionId, customProperties: customProperties);
        Assert.Contains("CustomProperties", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var isEvaluationMode = IsEvaluationMode();
        if (!isEvaluationMode) Assert.Equal("SessionValue", ppt.DocumentProperties["SessionKey"]?.ToString());
    }

    #endregion
}