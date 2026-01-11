using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordMailMergeTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordMailMergeToolTests : WordTestBase
{
    private readonly WordMailMergeTool _tool = new();

    #region File I/O Smoke Tests

    [Fact]
    public void PerformMailMerge_ShouldMergeDataAndPersistToFile()
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
        _tool.Execute(templatePath: templatePath, outputPath: outputPath, data: data.ToJsonString());
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("John", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("123 Main St", text, StringComparison.OrdinalIgnoreCase);
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
            new JsonObject { ["name"] = "Bob", ["id"] = "002" }
        };
        var result = _tool.Execute(templatePath: templatePath, outputPath: outputPath,
            dataArray: dataArray.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.Contains("Records processed: 2", result);

        var dir = Path.GetDirectoryName(outputPath)!;
        var baseName = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath);

        var file1 = Path.Combine(dir, $"{baseName}_1{ext}");
        var file2 = Path.Combine(dir, $"{baseName}_2{ext}");
        Assert.True(File.Exists(file1));
        Assert.True(File.Exists(file2));
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
        var data = new JsonObject { ["name"] = "TestUser" };
        var result = _tool.Execute(templatePath: templatePath, outputPath: outputPath, data: data.ToJsonString(),
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
        var result = _tool.Execute(templatePath: templatePath, outputPath: outputPath,
            dataArray: dataArray.ToJsonString());
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        Assert.Contains("SingleUser", resultDoc.GetText());
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("REMOVEUNUSEDFIELDS")]
    [InlineData("RemoveUnusedFields")]
    [InlineData("removeunusedfields")]
    public void CleanupOptions_ShouldBeCaseInsensitive(string option)
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
        var result = _tool.Execute(templatePath: templatePath, outputPath: outputPath, data: data.ToJsonString(),
            cleanupOptions: option);
        Assert.StartsWith("Mail merge completed successfully", result);
        Assert.True(File.Exists(outputPath));
    }

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
            _tool.Execute(templatePath: templatePath, outputPath: outputPath, data: data.ToJsonString(),
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
            _tool.Execute(templatePath: templatePath, outputPath: outputPath));
        Assert.Contains("must be provided", ex.Message);
    }

    [Fact]
    public void PerformMailMerge_WithMissingTemplateAndSession_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_missing_template_output.docx");
        var data = new JsonObject { ["name"] = "Test" };
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("execute", null, null, outputPath, data.ToJsonString()));
        Assert.Contains("Either templatePath or sessionId must be provided", ex.Message);
    }

    #endregion

    #region Session Management

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
    public void PerformMailMerge_WithInvalidSessionId_ShouldThrowException()
    {
        var tool = new WordMailMergeTool(SessionManager);

        var outputPath = CreateTestFilePath("test_invalid_session_output.docx");
        var data = new JsonObject { ["name"] = "Test" };

        Assert.ThrowsAny<Exception>(() =>
            tool.Execute(sessionId: "invalid_session_id", outputPath: outputPath, data: data.ToJsonString()));
    }

    #endregion
}
