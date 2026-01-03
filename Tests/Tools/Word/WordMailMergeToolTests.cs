using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordMailMergeToolTests : WordTestBase
{
    private readonly WordMailMergeTool _tool = new();

    #region General Tests

    [Fact]
    public void PerformMailMerge_ShouldMergeData()
    {
        var templatePath = CreateWordDocument("test_mail_merge_template.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", your address is ");
        builder.InsertField("MERGEFIELD address", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_output.docx");
        var data = new JsonObject
        {
            ["name"] = "John",
            ["address"] = "123 Main St"
        };
        _tool.Execute(templatePath, outputPath, data.ToJsonString());
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("John", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("123 Main St", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PerformMailMerge_WithMultipleFields_ShouldMergeAllFields()
    {
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
        var data = new JsonObject
        {
            ["firstName"] = "Jane",
            ["lastName"] = "Doe",
            ["orderNumber"] = "12345",
            ["city"] = "New York",
            ["country"] = "USA"
        };
        _tool.Execute(templatePath, outputPath, data.ToJsonString());
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Jane", text);
        Assert.Contains("Doe", text);
        Assert.Contains("12345", text);
        Assert.Contains("New York", text);
        Assert.Contains("USA", text);
    }

    [Fact]
    public void PerformMailMerge_WithEmptyValues_ShouldHandleEmptyFields()
    {
        var templatePath = CreateWordDocument("test_mail_merge_empty.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", Phone: ");
        builder.InsertField("MERGEFIELD phone", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_empty_output.docx");
        var data = new JsonObject
        {
            ["name"] = "TestUser",
            ["phone"] = "" // Empty value
        };
        _tool.Execute(templatePath, outputPath, data.ToJsonString());
        Assert.True(File.Exists(outputPath), "Output file should be created");
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("TestUser", text);
    }

    [Fact]
    public void PerformMailMerge_WithSpecialCharacters_ShouldHandleSpecialChars()
    {
        var templatePath = CreateWordDocument("test_mail_merge_special.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Company: ");
        builder.InsertField("MERGEFIELD company", "");
        builder.Write(", Email: ");
        builder.InsertField("MERGEFIELD email", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_special_output.docx");
        var data = new JsonObject
        {
            ["company"] = "Test & Co. <Ltd>",
            ["email"] = "test@example.com"
        };
        _tool.Execute(templatePath, outputPath, data.ToJsonString());
        Assert.True(File.Exists(outputPath), "Output file should be created");
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("test@example.com", text);
    }

    [Fact]
    public void PerformMailMerge_WithDataArray_ShouldGenerateMultipleFiles()
    {
        var templatePath = CreateWordDocument("test_mail_merge_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", your ID is ");
        builder.InsertField("MERGEFIELD id", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_array_output.docx");
        var dataArray = new JsonArray
        {
            new JsonObject { ["name"] = "Alice", ["id"] = "001" },
            new JsonObject { ["name"] = "Bob", ["id"] = "002" },
            new JsonObject { ["name"] = "Charlie", ["id"] = "003" }
        };
        var result = _tool.Execute(templatePath, outputPath, dataArray: dataArray.ToJsonString());
        Assert.Contains("multiple records", result);
        Assert.Contains("Records processed: 3", result);

        // Check individual files exist
        var dir = Path.GetDirectoryName(outputPath)!;
        var baseName = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath);

        for (var i = 1; i <= 3; i++)
        {
            var expectedFile = Path.Combine(dir, $"{baseName}_{i}{ext}");
            Assert.True(File.Exists(expectedFile), $"File {expectedFile} should exist");
        }

        // Verify content of first file
        var firstFile = Path.Combine(dir, $"{baseName}_1{ext}");
        var resultDoc = new Document(firstFile);
        var text = resultDoc.GetText();
        Assert.Contains("Alice", text);
        Assert.Contains("001", text);
    }

    [Fact]
    public void PerformMailMerge_WithCleanupOptions_ShouldApplyCleanup()
    {
        // Arrange - Create template with extra merge field that won't be populated
        var templatePath = CreateWordDocument("test_mail_merge_cleanup.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Writeln();
        builder.Write("Unused: ");
        builder.InsertField("MERGEFIELD unusedField", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_cleanup_output.docx");
        var data = new JsonObject
        {
            ["name"] = "TestUser"
        };
        var result = _tool.Execute(templatePath, outputPath, data.ToJsonString(),
            cleanupOptions: "removeUnusedFields,removeEmptyParagraphs");
        Assert.Contains("Cleanup applied", result);
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("TestUser", text);
        // The unused field should be removed
        Assert.DoesNotContain("unusedField", text);
    }

    [Fact]
    public void PerformMailMerge_WithSingleRecordInDataArray_ShouldNotAddSuffix()
    {
        var templatePath = CreateWordDocument("test_mail_merge_single_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_single_array_output.docx");
        var dataArray = new JsonArray
        {
            new JsonObject { ["name"] = "SingleUser" }
        };
        var result = _tool.Execute(templatePath, outputPath, dataArray: dataArray.ToJsonString());
        Assert.Contains("Records processed: 1", result);
        // Single record should use the original output path
        Assert.True(File.Exists(outputPath), "Original output path should be used for single record");
        var resultDoc = new Document(outputPath);
        Assert.Contains("SingleUser", resultDoc.GetText());
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void PerformMailMerge_WithBothDataAndDataArray_ShouldReturnError()
    {
        var templatePath = CreateWordDocument("test_mail_merge_error.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_error_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var dataArray = new JsonArray { new JsonObject { ["name"] = "Test2" } };
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(templatePath, outputPath, data.ToJsonString(), dataArray.ToJsonString()));
        Assert.Contains("Cannot specify both", ex.Message);
    }

    [Fact]
    public void PerformMailMerge_WithNoData_ShouldReturnError()
    {
        var templatePath = CreateWordDocument("test_mail_merge_nodata.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_nodata_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(templatePath, outputPath));
        Assert.Contains("must be provided", ex.Message);
    }

    #endregion

    // Note: This tool does not support session, so no Session ID Tests region
}