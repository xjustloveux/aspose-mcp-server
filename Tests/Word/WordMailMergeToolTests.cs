using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordMailMergeToolTests : WordTestBase
{
    private readonly WordMailMergeTool _tool = new();

    [Fact]
    public async Task PerformMailMerge_ShouldMergeData()
    {
        // Arrange
        var templatePath = CreateWordDocument("test_mail_merge_template.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", your address is ");
        builder.InsertField("MERGEFIELD address", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_output.docx");
        var arguments = new JsonObject
        {
            ["templatePath"] = templatePath,
            ["outputPath"] = outputPath,
            ["data"] = new JsonObject
            {
                ["name"] = "John",
                ["address"] = "123 Main St"
            }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("John", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("123 Main St", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task PerformMailMerge_WithMultipleFields_ShouldMergeAllFields()
    {
        // Arrange
        var templatePath = CreateWordDocument("test_mail_merge_multi.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD firstName", "");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD lastName", "");
        builder.Writeln(",");
        builder.Write("Your order #");
        builder.InsertField("MERGEFIELD orderNumber", "");
        builder.Write(" will be shipped to ");
        builder.InsertField("MERGEFIELD city", "");
        builder.Write(", ");
        builder.InsertField("MERGEFIELD country", "");
        builder.Write(".");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_multi_output.docx");
        var arguments = new JsonObject
        {
            ["templatePath"] = templatePath,
            ["outputPath"] = outputPath,
            ["data"] = new JsonObject
            {
                ["firstName"] = "Jane",
                ["lastName"] = "Doe",
                ["orderNumber"] = "12345",
                ["city"] = "New York",
                ["country"] = "USA"
            }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Jane", text);
        Assert.Contains("Doe", text);
        Assert.Contains("12345", text);
        Assert.Contains("New York", text);
        Assert.Contains("USA", text);
    }

    [Fact]
    public async Task PerformMailMerge_WithEmptyValues_ShouldHandleEmptyFields()
    {
        // Arrange
        var templatePath = CreateWordDocument("test_mail_merge_empty.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", Phone: ");
        builder.InsertField("MERGEFIELD phone", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_empty_output.docx");
        var arguments = new JsonObject
        {
            ["templatePath"] = templatePath,
            ["outputPath"] = outputPath,
            ["data"] = new JsonObject
            {
                ["name"] = "TestUser",
                ["phone"] = "" // Empty value
            }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output file should be created");
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("TestUser", text);
    }

    [Fact]
    public async Task PerformMailMerge_WithSpecialCharacters_ShouldHandleSpecialChars()
    {
        // Arrange
        var templatePath = CreateWordDocument("test_mail_merge_special.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Company: ");
        builder.InsertField("MERGEFIELD company", "");
        builder.Write(", Email: ");
        builder.InsertField("MERGEFIELD email", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_special_output.docx");
        var arguments = new JsonObject
        {
            ["templatePath"] = templatePath,
            ["outputPath"] = outputPath,
            ["data"] = new JsonObject
            {
                ["company"] = "Test & Co. <Ltd>",
                ["email"] = "test@example.com"
            }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output file should be created");
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("test@example.com", text);
    }
}