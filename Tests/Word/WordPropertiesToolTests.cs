using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordPropertiesToolTests : WordTestBase
{
    private readonly WordPropertiesTool _tool = new();

    [Fact]
    public async Task GetProperties_ShouldReturnProperties()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_properties.docx");
        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Properties", result, StringComparison.OrdinalIgnoreCase);
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
}