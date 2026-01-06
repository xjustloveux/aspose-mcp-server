using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordMailMergeToolTests : WordTestBase
{
    private readonly WordMailMergeTool _tool = new();

    #region General

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
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
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
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
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
            ["phone"] = ""
        };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
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
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
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
        var result = _tool.Execute(templatePath, outputPath: outputPath, dataArray: dataArray.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.Contains("Records processed: 3", result);

        var dir = Path.GetDirectoryName(outputPath)!;
        var baseName = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath);

        for (var i = 1; i <= 3; i++)
        {
            var expectedFile = Path.Combine(dir, $"{baseName}_{i}{ext}");
            Assert.True(File.Exists(expectedFile), $"File {expectedFile} should exist");
        }

        var firstFile = Path.Combine(dir, $"{baseName}_1{ext}");
        var resultDoc = new Document(firstFile);
        var text = resultDoc.GetText();
        Assert.Contains("Alice", text);
        Assert.Contains("001", text);
    }

    [Fact]
    public void PerformMailMerge_WithCleanupOptions_ShouldApplyCleanup()
    {
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
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: "removeUnusedFields,removeEmptyParagraphs");
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("TestUser", text);
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
        var result = _tool.Execute(templatePath, outputPath: outputPath, dataArray: dataArray.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        Assert.Contains("SingleUser", resultDoc.GetText());
    }

    [Fact]
    public void PerformMailMerge_WithNullValues_ShouldHandleNullFields()
    {
        var templatePath = CreateWordDocument("test_mail_merge_null.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", Value: ");
        builder.InsertField("MERGEFIELD value", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_null_output.docx");
        var data = new JsonObject
        {
            ["name"] = "TestUser",
            ["value"] = null
        };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("TestUser", text);
    }

    [Fact]
    public void PerformMailMerge_WithUnicodeCharacters_ShouldHandleUnicode()
    {
        var templatePath = CreateWordDocument("test_mail_merge_unicode.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", City: ");
        builder.InsertField("MERGEFIELD city", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_unicode_output.docx");
        var data = new JsonObject
        {
            ["name"] = "張三",
            ["city"] = "東京"
        };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("張三", text);
        Assert.Contains("東京", text);
    }

    [Fact]
    public void PerformMailMerge_WithTemplateWithoutMergeFields_ShouldSucceed()
    {
        var templatePath = CreateWordDocumentWithContent("test_mail_merge_no_fields.docx",
            "This is plain text without merge fields.");
        var outputPath = CreateTestFilePath("test_mail_merge_no_fields_output.docx");
        var data = new JsonObject
        {
            ["name"] = "Test"
        };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithRemoveUnusedRegions_ShouldApplyCleanup()
    {
        var templatePath = CreateWordDocument("test_mail_merge_regions.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_regions_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: "removeUnusedRegions");
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithRemoveContainingFields_ShouldApplyCleanup()
    {
        var templatePath = CreateWordDocument("test_mail_merge_containing.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Writeln();
        builder.Write("Optional: ");
        builder.InsertField("MERGEFIELD optional", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_containing_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: "removeContainingFields");
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithRemoveStaticFields_ShouldApplyCleanup()
    {
        var templatePath = CreateWordDocument("test_mail_merge_static.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.InsertField("PAGE", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_static_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: "removeStaticFields");
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithMultipleCleanupOptions_ShouldApplyAll()
    {
        var templatePath = CreateWordDocument("test_mail_merge_multi_cleanup.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Writeln();
        builder.Write("Unused: ");
        builder.InsertField("MERGEFIELD unused", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_multi_cleanup_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: "removeUnusedFields,removeEmptyParagraphs,removeContainingFields");
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithInvalidCleanupOption_ShouldIgnoreInvalid()
    {
        var templatePath = CreateWordDocument("test_mail_merge_invalid_cleanup.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_invalid_cleanup_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: "invalidOption,removeUnusedFields");
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithDataArrayContainingMultipleFields_ShouldMergeAll()
    {
        var templatePath = CreateWordDocument("test_mail_merge_array_multi.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", Age: ");
        builder.InsertField("MERGEFIELD age", "");
        builder.Write(", City: ");
        builder.InsertField("MERGEFIELD city", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_array_multi_output.docx");
        var dataArray = new JsonArray
        {
            new JsonObject { ["name"] = "Alice", ["age"] = "25", ["city"] = "NYC" },
            new JsonObject { ["name"] = "Bob", ["age"] = "30", ["city"] = "LA" }
        };
        var result = _tool.Execute(templatePath, outputPath: outputPath, dataArray: dataArray.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);

        var dir = Path.GetDirectoryName(outputPath)!;
        var baseName = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath);

        var file1 = Path.Combine(dir, $"{baseName}_1{ext}");
        var file2 = Path.Combine(dir, $"{baseName}_2{ext}");

        var doc1 = new Document(file1);
        Assert.Contains("Alice", doc1.GetText());
        Assert.Contains("25", doc1.GetText());
        Assert.Contains("NYC", doc1.GetText());

        var doc2 = new Document(file2);
        Assert.Contains("Bob", doc2.GetText());
        Assert.Contains("30", doc2.GetText());
        Assert.Contains("LA", doc2.GetText());
    }

    [Fact]
    public void PerformMailMerge_WithMissingFieldInData_ShouldLeaveFieldUnpopulated()
    {
        var templatePath = CreateWordDocument("test_mail_merge_missing_field.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(", Phone: ");
        builder.InsertField("MERGEFIELD phone", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_missing_field_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Test", text);
    }

    [Fact]
    public void PerformMailMerge_WithEmptyStringData_ShouldMergeEmptyValue()
    {
        var templatePath = CreateWordDocument("test_mail_merge_empty_string.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: [");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write("]");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_empty_string_output.docx");
        var data = new JsonObject { ["name"] = "" };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithVeryLongFieldValue_ShouldHandleLongText()
    {
        var templatePath = CreateWordDocument("test_mail_merge_long.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Content: ");
        builder.InsertField("MERGEFIELD content", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_long_output.docx");
        var longText = new string('A', 10000);
        var data = new JsonObject { ["content"] = longText };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains(longText, text);
    }

    [Fact]
    public void PerformMailMerge_WithNumericValues_ShouldConvertToString()
    {
        var templatePath = CreateWordDocument("test_mail_merge_numeric.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Amount: ");
        builder.InsertField("MERGEFIELD amount", "");
        builder.Write(", Count: ");
        builder.InsertField("MERGEFIELD count", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_numeric_output.docx");
        var data = new JsonObject
        {
            ["amount"] = 123.45,
            ["count"] = 100
        };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("123.45", text);
        Assert.Contains("100", text);
    }

    [Fact]
    public void PerformMailMerge_WithEmptyDataObject_ShouldSucceed()
    {
        var templatePath = CreateWordDocumentWithContent("test_mail_merge_empty_obj.docx", "Plain text");
        var outputPath = CreateTestFilePath("test_mail_merge_empty_obj_output.docx");
        var data = new JsonObject();
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithDataArrayContainingNullRecord_ShouldSkipNull()
    {
        var templatePath = CreateWordDocument("test_mail_merge_null_record.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_null_record_output.docx");
        var dataArray = new JsonArray
        {
            new JsonObject { ["name"] = "Alice" },
            null,
            new JsonObject { ["name"] = "Bob" }
        };
        var result = _tool.Execute(templatePath, outputPath: outputPath, dataArray: dataArray.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);
    }

    [Fact]
    public void PerformMailMerge_WithBooleanValues_ShouldConvertToString()
    {
        var templatePath = CreateWordDocument("test_mail_merge_bool.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Active: ");
        builder.InsertField("MERGEFIELD active", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_bool_output.docx");
        var data = new JsonObject { ["active"] = true };
        _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString());
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("true", text);
    }

    [Theory]
    [InlineData("REMOVEUNUSEDFIELDS")]
    [InlineData("RemoveUnusedFields")]
    [InlineData("removeunusedfields")]
    [InlineData("removeUnusedFields")]
    public void CleanupOptions_ShouldBeCaseInsensitive_RemoveUnusedFields(string option)
    {
        var templatePath = CreateWordDocument($"test_cleanup_case_{option}.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(" Unused: ");
        builder.InsertField("MERGEFIELD unused", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath($"test_cleanup_case_{option}_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: option);
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("REMOVEEMPTYPARAGRAPHS")]
    [InlineData("RemoveEmptyParagraphs")]
    [InlineData("removeemptyparagraphs")]
    [InlineData("removeEmptyParagraphs")]
    public void CleanupOptions_ShouldBeCaseInsensitive_RemoveEmptyParagraphs(string option)
    {
        var templatePath = CreateWordDocument($"test_cleanup_paragraphs_{option}.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath($"test_cleanup_paragraphs_{option}_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: option);
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("REMOVEUNUSEDREGIONS")]
    [InlineData("RemoveUnusedRegions")]
    [InlineData("removeunusedregions")]
    public void CleanupOptions_ShouldBeCaseInsensitive_RemoveUnusedRegions(string option)
    {
        var templatePath = CreateWordDocument($"test_cleanup_regions_{option}.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath($"test_cleanup_regions_{option}_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: option);
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("REMOVECONTAININGFIELDS")]
    [InlineData("RemoveContainingFields")]
    [InlineData("removecontainingfields")]
    public void CleanupOptions_ShouldBeCaseInsensitive_RemoveContainingFields(string option)
    {
        var templatePath = CreateWordDocument($"test_cleanup_containing_{option}.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath($"test_cleanup_containing_{option}_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: option);
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("REMOVESTATICFIELDS")]
    [InlineData("RemoveStaticFields")]
    [InlineData("removestaticfields")]
    public void CleanupOptions_ShouldBeCaseInsensitive_RemoveStaticFields(string option)
    {
        var templatePath = CreateWordDocument($"test_cleanup_static_{option}.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath($"test_cleanup_static_{option}_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var result = _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: option);
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Exception

    [Fact]
    public void PerformMailMerge_WithBothDataAndDataArray_ShouldThrowException()
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
            _tool.Execute(templatePath, outputPath: outputPath, data: data.ToJsonString(),
                dataArray: dataArray.ToJsonString()));
        Assert.Contains("Cannot specify both", ex.Message);
    }

    [Fact]
    public void PerformMailMerge_WithNoData_ShouldThrowException()
    {
        var templatePath = CreateWordDocument("test_mail_merge_nodata.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_nodata_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(templatePath, outputPath: outputPath));
        Assert.Contains("must be provided", ex.Message);
    }

    [Fact]
    public void PerformMailMerge_WithInvalidDataJson_ShouldThrowException()
    {
        var templatePath = CreateWordDocument("test_mail_merge_invalid_data.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_invalid_data_output.docx");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute(templatePath, outputPath: outputPath, data: "not valid json"));
    }

    [Fact]
    public void PerformMailMerge_WithInvalidDataArrayJson_ShouldThrowException()
    {
        var templatePath = CreateWordDocument("test_mail_merge_invalid_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_invalid_array_output.docx");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute(templatePath, outputPath: outputPath, dataArray: "not valid json"));
    }

    [Fact]
    public void PerformMailMerge_WithEmptyDataArray_ShouldThrowException()
    {
        var templatePath = CreateWordDocument("test_mail_merge_empty_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var outputPath = CreateTestFilePath("test_mail_merge_empty_array_output.docx");
        var dataArray = new JsonArray();
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(templatePath, outputPath: outputPath, dataArray: dataArray.ToJsonString()));
        Assert.Contains("No data provided", ex.Message);
    }

    [Fact]
    public void PerformMailMerge_WithNonExistentTemplate_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_mail_merge_nonexistent_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("nonexistent_template.docx", outputPath: outputPath, data: data.ToJsonString()));
    }

    [Fact]
    public void PerformMailMerge_WithMissingOutputPath_ShouldThrowException()
    {
        var templatePath = CreateWordDocumentWithContent("test_missing_output.docx", "Test");
        var data = new JsonObject { ["name"] = "Test" };
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(templatePath, outputPath: null, data: data.ToJsonString()));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void PerformMailMerge_WithMissingTemplateAndSession_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_missing_template_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(null, null, outputPath, data.ToJsonString()));
        Assert.Contains("Either templatePath or sessionId must be provided", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void PerformMailMerge_WithSessionId_ShouldMergeFromSession()
    {
        var tool = new WordMailMergeTool(SessionManager);

        var templatePath = CreateWordDocument("test_session_mail_merge.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write("!");
        doc.Save(templatePath);

        var sessionId = SessionManager.OpenDocument(templatePath);
        var outputPath = CreateTestFilePath("test_session_mail_merge_output.docx");
        var data = new JsonObject { ["name"] = "SessionUser" };

        var result = tool.Execute(sessionId: sessionId, outputPath: outputPath, data: data.ToJsonString());

        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        Assert.Contains("SessionUser", resultDoc.GetText());
    }

    [Fact]
    public void PerformMailMerge_WithSessionId_ShouldNotModifyOriginalSession()
    {
        var tool = new WordMailMergeTool(SessionManager);

        var templatePath = CreateWordDocument("test_session_original.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var sessionId = SessionManager.OpenDocument(templatePath);
        var originalDoc = SessionManager.GetDocument<Document>(sessionId);
        var originalText = originalDoc.GetText();

        var outputPath = CreateTestFilePath("test_session_original_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        tool.Execute(sessionId: sessionId, outputPath: outputPath, data: data.ToJsonString());

        var sessionDocAfter = SessionManager.GetDocument<Document>(sessionId);
        var textAfter = sessionDocAfter.GetText();
        Assert.Equal(originalText, textAfter);
    }

    [Fact]
    public void PerformMailMerge_WithSessionIdAndDataArray_ShouldGenerateMultipleFiles()
    {
        var tool = new WordMailMergeTool(SessionManager);

        var templatePath = CreateWordDocument("test_session_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var sessionId = SessionManager.OpenDocument(templatePath);
        var outputPath = CreateTestFilePath("test_session_array_output.docx");
        var dataArray = new JsonArray
        {
            new JsonObject { ["name"] = "Alice" },
            new JsonObject { ["name"] = "Bob" }
        };

        var result = tool.Execute(sessionId: sessionId, outputPath: outputPath, dataArray: dataArray.ToJsonString());

        Assert.StartsWith("Mail merge completed successfully", result);

        var dir = Path.GetDirectoryName(outputPath)!;
        var baseName = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath);

        var file1 = Path.Combine(dir, $"{baseName}_1{ext}");
        var file2 = Path.Combine(dir, $"{baseName}_2{ext}");

        Assert.True(File.Exists(file1));
        Assert.True(File.Exists(file2));
        Assert.Contains("Alice", new Document(file1).GetText());
        Assert.Contains("Bob", new Document(file2).GetText());
    }

    [Fact]
    public void PerformMailMerge_WithSessionIdAndCleanup_ShouldApplyCleanup()
    {
        var tool = new WordMailMergeTool(SessionManager);

        var templatePath = CreateWordDocument("test_session_cleanup.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Name: ");
        builder.InsertField("MERGEFIELD name", "");
        builder.Write(" Unused: ");
        builder.InsertField("MERGEFIELD unused", "");
        doc.Save(templatePath);

        var sessionId = SessionManager.OpenDocument(templatePath);
        var outputPath = CreateTestFilePath("test_session_cleanup_output.docx");
        var data = new JsonObject { ["name"] = "Test" };

        var result = tool.Execute(sessionId: sessionId, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: "removeUnusedFields");

        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void PerformMailMerge_WithInvalidSessionId_ShouldThrowException()
    {
        var tool = new WordMailMergeTool(SessionManager);

        var outputPath = CreateTestFilePath("test_invalid_session_output.docx");
        var data = new JsonObject { ["name"] = "Test" };

        Assert.ThrowsAny<Exception>(() =>
            tool.Execute(sessionId: "invalid_session_id", outputPath: outputPath, data: data.ToJsonString()));
    }

    [Fact]
    public void PerformMailMerge_WithSessionIdAndSingleRecordArray_ShouldNotAddSuffix()
    {
        var tool = new WordMailMergeTool(SessionManager);

        var templatePath = CreateWordDocument("test_session_single_array.docx");
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD name", "");
        doc.Save(templatePath);

        var sessionId = SessionManager.OpenDocument(templatePath);
        var outputPath = CreateTestFilePath("test_session_single_array_output.docx");
        var dataArray = new JsonArray
        {
            new JsonObject { ["name"] = "Single" }
        };

        var result = tool.Execute(sessionId: sessionId, outputPath: outputPath, dataArray: dataArray.ToJsonString());

        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Single", new Document(outputPath).GetText());
    }

    #endregion
}