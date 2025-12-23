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
}