using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Properties;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordPropertiesToolTests : WordTestBase
{
    private readonly WordPropertiesTool _tool = new();

    [Fact]
    public async Task GetProperties_ShouldReturnJsonFormat()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_properties.docx");
        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);

        // Verify it's valid JSON
        var json = JsonNode.Parse(result);
        Assert.NotNull(json);

        // Verify structure
        Assert.NotNull(json["builtInProperties"]);
        Assert.NotNull(json["statistics"]);
        Assert.NotNull(json["statistics"]?["wordCount"]);
        Assert.NotNull(json["statistics"]?["pageCount"]);
    }

    [Fact]
    public async Task SetProperties_ShouldSetProperties()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_properties.docx");
        var outputPath = CreateTestFilePath("test_set_properties_output.docx");
        var arguments = CreateArguments("set", docPath, outputPath);
        arguments["title"] = "Test Document";
        arguments["author"] = "Test Author";
        arguments["subject"] = "Test Subject";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal("Test Document", doc.BuiltInDocumentProperties.Title);
        Assert.Equal("Test Author", doc.BuiltInDocumentProperties.Author);
        Assert.Equal("Test Subject", doc.BuiltInDocumentProperties.Subject);
    }

    [Fact]
    public async Task SetProperties_WithAllBuiltInProperties_ShouldSetAll()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_all_properties.docx");
        var outputPath = CreateTestFilePath("test_set_all_properties_output.docx");
        var arguments = CreateArguments("set", docPath, outputPath);
        arguments["title"] = "Full Title";
        arguments["author"] = "Full Author";
        arguments["subject"] = "Full Subject";
        arguments["keywords"] = "keyword1, keyword2";
        arguments["comments"] = "Test comments";
        arguments["category"] = "Test Category";
        arguments["company"] = "Test Company";
        arguments["manager"] = "Test Manager";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var props = doc.BuiltInDocumentProperties;
        Assert.Equal("Full Title", props.Title);
        Assert.Equal("Full Author", props.Author);
        Assert.Equal("Full Subject", props.Subject);
        Assert.Equal("keyword1, keyword2", props.Keywords);
        Assert.Equal("Test comments", props.Comments);
        Assert.Equal("Test Category", props.Category);
        Assert.Equal("Test Company", props.Company);
        Assert.Equal("Test Manager", props.Manager);
    }

    [Fact]
    public async Task SetProperties_WithCustomStringProperty_ShouldAddAsString()
    {
        // Arrange
        var docPath = CreateWordDocument("test_custom_string.docx");
        var outputPath = CreateTestFilePath("test_custom_string_output.docx");
        var arguments = CreateArguments("set", docPath, outputPath);
        arguments["customProperties"] = new JsonObject
        {
            ["ProjectName"] = "My Project",
            ["Version"] = "1.0.0"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal("My Project", doc.CustomDocumentProperties["ProjectName"]?.Value?.ToString());
        Assert.Equal("1.0.0", doc.CustomDocumentProperties["Version"]?.Value?.ToString());
    }

    [Fact]
    public async Task SetProperties_WithCustomIntegerProperty_ShouldAddAsInteger()
    {
        // Arrange
        var docPath = CreateWordDocument("test_custom_int.docx");
        var outputPath = CreateTestFilePath("test_custom_int_output.docx");
        var arguments = CreateArguments("set", docPath, outputPath);
        arguments["customProperties"] = new JsonObject
        {
            ["RevisionNumber"] = 42,
            ["Priority"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var revProp = doc.CustomDocumentProperties["RevisionNumber"];
        Assert.NotNull(revProp);
        Assert.Equal(PropertyType.Number, revProp.Type);
        Assert.Equal(42, revProp.ToInt());
    }

    [Fact]
    public async Task SetProperties_WithCustomBooleanProperty_ShouldAddAsBoolean()
    {
        // Arrange
        var docPath = CreateWordDocument("test_custom_bool.docx");
        var outputPath = CreateTestFilePath("test_custom_bool_output.docx");
        var arguments = CreateArguments("set", docPath, outputPath);
        arguments["customProperties"] = new JsonObject
        {
            ["IsApproved"] = true,
            ["IsArchived"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var approvedProp = doc.CustomDocumentProperties["IsApproved"];
        Assert.NotNull(approvedProp);
        Assert.Equal(PropertyType.Boolean, approvedProp.Type);
        Assert.True(approvedProp.ToBool());

        var archivedProp = doc.CustomDocumentProperties["IsArchived"];
        Assert.NotNull(archivedProp);
        Assert.False(archivedProp.ToBool());
    }

    [Fact]
    public async Task SetProperties_WithCustomDoubleProperty_ShouldAddAsDouble()
    {
        // Arrange
        var docPath = CreateWordDocument("test_custom_double.docx");
        var outputPath = CreateTestFilePath("test_custom_double_output.docx");
        var arguments = CreateArguments("set", docPath, outputPath);
        arguments["customProperties"] = new JsonObject
        {
            ["Price"] = 99.99,
            ["Discount"] = 0.15
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var priceProp = doc.CustomDocumentProperties["Price"];
        Assert.NotNull(priceProp);
        Assert.Equal(PropertyType.Double, priceProp.Type);
        Assert.Equal(99.99, priceProp.ToDouble(), 2);
    }

    [Fact]
    public async Task SetProperties_WithExistingCustomProperty_ShouldUpdate()
    {
        // Arrange - First create a document with a custom property
        var docPath = CreateWordDocument("test_update_custom.docx");
        var doc = new Document(docPath);
        doc.CustomDocumentProperties.Add("Status", "Draft");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_custom_output.docx");
        var arguments = CreateArguments("set", docPath, outputPath);
        arguments["customProperties"] = new JsonObject
        {
            ["Status"] = "Published"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.Equal("Published", resultDoc.CustomDocumentProperties["Status"]?.Value?.ToString());
    }

    [Fact]
    public async Task GetProperties_WithCustomProperties_ShouldReturnInJson()
    {
        // Arrange - Create a document with custom properties of different types
        var docPath = CreateWordDocument("test_get_custom_types.docx");
        var doc = new Document(docPath);
        doc.CustomDocumentProperties.Add("StringProp", "Hello");
        doc.CustomDocumentProperties.Add("IntProp", 123);
        doc.CustomDocumentProperties.Add("BoolProp", true);
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert - Verify JSON structure
        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.NotNull(json["customProperties"]);

        var customProps = json["customProperties"]!;
        Assert.NotNull(customProps["StringProp"]);
        Assert.NotNull(customProps["IntProp"]);
        Assert.NotNull(customProps["BoolProp"]);

        // Verify type info is included
        Assert.Equal("String", customProps["StringProp"]?["type"]?.GetValue<string>());
        Assert.Equal("Number", customProps["IntProp"]?["type"]?.GetValue<string>());
        Assert.Equal("Boolean", customProps["BoolProp"]?["type"]?.GetValue<string>());
    }

    [Fact]
    public async Task GetProperties_ShouldIncludeAllStatistics()
    {
        // Arrange
        var docPath =
            CreateWordDocumentWithContent("test_get_statistics.docx", "Hello World. This is a test document.");
        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonNode.Parse(result);
        Assert.NotNull(json);

        var stats = json["statistics"];
        Assert.NotNull(stats);
        Assert.NotNull(stats["wordCount"]);
        Assert.NotNull(stats["characterCount"]);
        Assert.NotNull(stats["pageCount"]);
        Assert.NotNull(stats["paragraphCount"]);
        Assert.NotNull(stats["lineCount"]);

        // Word count should be > 0 for non-empty document
        Assert.True(stats["wordCount"]?.GetValue<int>() > 0);
    }
}