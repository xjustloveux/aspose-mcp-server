using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Properties;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordPropertiesToolTests : WordTestBase
{
    private readonly WordPropertiesTool _tool;

    public WordPropertiesToolTests()
    {
        _tool = new WordPropertiesTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void GetProperties_ShouldReturnJsonFormat()
    {
        var docPath = CreateWordDocument("test_get_properties.docx");
        var result = _tool.Execute("get", docPath);
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
    public void SetProperties_ShouldSetProperties()
    {
        var docPath = CreateWordDocument("test_set_properties.docx");
        var outputPath = CreateTestFilePath("test_set_properties_output.docx");
        _tool.Execute("set", docPath, outputPath: outputPath,
            title: "Test Document", author: "Test Author", subject: "Test Subject");
        var doc = new Document(outputPath);
        Assert.Equal("Test Document", doc.BuiltInDocumentProperties.Title);
        Assert.Equal("Test Author", doc.BuiltInDocumentProperties.Author);
        Assert.Equal("Test Subject", doc.BuiltInDocumentProperties.Subject);
    }

    [Fact]
    public void SetProperties_WithAllBuiltInProperties_ShouldSetAll()
    {
        var docPath = CreateWordDocument("test_set_all_properties.docx");
        var outputPath = CreateTestFilePath("test_set_all_properties_output.docx");
        _tool.Execute("set", docPath, outputPath: outputPath,
            title: "Full Title", author: "Full Author", subject: "Full Subject",
            keywords: "keyword1, keyword2", comments: "Test comments",
            category: "Test Category", company: "Test Company", manager: "Test Manager");
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
    public void SetProperties_WithCustomStringProperty_ShouldAddAsString()
    {
        var docPath = CreateWordDocument("test_custom_string.docx");
        var outputPath = CreateTestFilePath("test_custom_string_output.docx");
        var customProps = new JsonObject
        {
            ["ProjectName"] = "My Project",
            ["Version"] = "1.0.0"
        };
        _tool.Execute("set", docPath, outputPath: outputPath, customProperties: customProps.ToJsonString());
        var doc = new Document(outputPath);
        Assert.Equal("My Project", doc.CustomDocumentProperties["ProjectName"]?.Value?.ToString());
        Assert.Equal("1.0.0", doc.CustomDocumentProperties["Version"]?.Value?.ToString());
    }

    [Fact]
    public void SetProperties_WithCustomIntegerProperty_ShouldAddAsInteger()
    {
        var docPath = CreateWordDocument("test_custom_int.docx");
        var outputPath = CreateTestFilePath("test_custom_int_output.docx");
        var customProps = new JsonObject
        {
            ["RevisionNumber"] = 42,
            ["Priority"] = 1
        };
        _tool.Execute("set", docPath, outputPath: outputPath, customProperties: customProps.ToJsonString());
        var doc = new Document(outputPath);
        var revProp = doc.CustomDocumentProperties["RevisionNumber"];
        Assert.NotNull(revProp);
        Assert.Equal(PropertyType.Number, revProp.Type);
        Assert.Equal(42, revProp.ToInt());
    }

    [Fact]
    public void SetProperties_WithCustomBooleanProperty_ShouldAddAsBoolean()
    {
        var docPath = CreateWordDocument("test_custom_bool.docx");
        var outputPath = CreateTestFilePath("test_custom_bool_output.docx");
        var customProps = new JsonObject
        {
            ["IsApproved"] = true,
            ["IsArchived"] = false
        };
        _tool.Execute("set", docPath, outputPath: outputPath, customProperties: customProps.ToJsonString());
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
    public void SetProperties_WithCustomDoubleProperty_ShouldAddAsDouble()
    {
        var docPath = CreateWordDocument("test_custom_double.docx");
        var outputPath = CreateTestFilePath("test_custom_double_output.docx");
        var customProps = new JsonObject
        {
            ["Price"] = 99.99,
            ["Discount"] = 0.15
        };
        _tool.Execute("set", docPath, outputPath: outputPath, customProperties: customProps.ToJsonString());
        var doc = new Document(outputPath);
        var priceProp = doc.CustomDocumentProperties["Price"];
        Assert.NotNull(priceProp);
        Assert.Equal(PropertyType.Double, priceProp.Type);
        Assert.Equal(99.99, priceProp.ToDouble(), 2);
    }

    [Fact]
    public void SetProperties_WithExistingCustomProperty_ShouldUpdate()
    {
        // Arrange - First create a document with a custom property
        var docPath = CreateWordDocument("test_update_custom.docx");
        var doc = new Document(docPath);
        doc.CustomDocumentProperties.Add("Status", "Draft");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_custom_output.docx");
        var customProps = new JsonObject
        {
            ["Status"] = "Published"
        };
        _tool.Execute("set", docPath, outputPath: outputPath, customProperties: customProps.ToJsonString());
        var resultDoc = new Document(outputPath);
        Assert.Equal("Published", resultDoc.CustomDocumentProperties["Status"]?.Value?.ToString());
    }

    [Fact]
    public void GetProperties_WithCustomProperties_ShouldReturnInJson()
    {
        // Arrange - Create a document with custom properties of different types
        var docPath = CreateWordDocument("test_get_custom_types.docx");
        var doc = new Document(docPath);
        doc.CustomDocumentProperties.Add("StringProp", "Hello");
        doc.CustomDocumentProperties.Add("IntProp", 123);
        doc.CustomDocumentProperties.Add("BoolProp", true);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);

        // Assert - Verify JSON structure
        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.NotNull(json["customProperties"]);

        var customPropsJson = json["customProperties"]!;
        Assert.NotNull(customPropsJson["StringProp"]);
        Assert.NotNull(customPropsJson["IntProp"]);
        Assert.NotNull(customPropsJson["BoolProp"]);

        // Verify type info is included
        Assert.Equal("String", customPropsJson["StringProp"]?["type"]?.GetValue<string>());
        Assert.Equal("Number", customPropsJson["IntProp"]?["type"]?.GetValue<string>());
        Assert.Equal("Boolean", customPropsJson["BoolProp"]?["type"]?.GetValue<string>());
    }

    [Fact]
    public void GetProperties_ShouldIncludeAllStatistics()
    {
        var docPath =
            CreateWordDocumentWithContent("test_get_statistics.docx", "Hello World. This is a test document.");
        var result = _tool.Execute("get", docPath);
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

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void SetProperties_WithInvalidCustomPropertiesJson_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_invalid_json.docx");
        var outputPath = CreateTestFilePath("test_invalid_json_output.docx");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("set", docPath, outputPath: outputPath, customProperties: "not valid json"));
    }

    [Fact]
    public void SetProperties_WithNoProperties_ShouldSucceedWithNoChanges()
    {
        var docPath = CreateWordDocument("test_no_properties.docx");
        var outputPath = CreateTestFilePath("test_no_properties_output.docx");

        // Act - No properties provided, should still succeed
        var result = _tool.Execute("set", docPath, outputPath: outputPath);

        // Assert - Should succeed without any changes
        Assert.Contains("properties", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetProperties_WithSessionId_ShouldReturnProperties()
    {
        var docPath = CreateWordDocument("test_session_get_props.docx");
        var doc = new Document(docPath) { BuiltInDocumentProperties = { Title = "Session Title" } };
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("Session Title", result);
    }

    [Fact]
    public void SetProperties_WithSessionId_ShouldSetPropertiesInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_props.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set", sessionId: sessionId,
            title: "In-Memory Title", author: "In-Memory Author");

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("In-Memory Title", sessionDoc.BuiltInDocumentProperties.Title);
        Assert.Equal("In-Memory Author", sessionDoc.BuiltInDocumentProperties.Author);
    }

    [Fact]
    public void SetProperties_WithSessionId_ShouldModifyCustomPropertiesInMemory()
    {
        var docPath = CreateWordDocument("test_session_custom_props.docx");
        var sessionId = OpenSession(docPath);
        var customProps = new JsonObject
        {
            ["SessionProp"] = "SessionValue"
        };
        _tool.Execute("set", sessionId: sessionId, customProperties: customProps.ToJsonString());

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("SessionValue", sessionDoc.CustomDocumentProperties["SessionProp"]?.Value?.ToString());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_props.docx");
        var doc1 = new Document(docPath1) { BuiltInDocumentProperties = { Title = "Path Title" } };
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_props.docx");
        var doc2 = new Document(docPath2) { BuiltInDocumentProperties = { Title = "Session Title" } };
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        var result = _tool.Execute("get", docPath1, sessionId);

        // Assert - should use sessionId, returning Session Title not Path Title
        Assert.Contains("Session Title", result);
        Assert.DoesNotContain("Path Title", result);
    }

    [Fact]
    public void GetProperties_WithSessionId_VerifyInMemoryChangesReflected()
    {
        var docPath = CreateWordDocument("test_session_verify.docx");
        var sessionId = OpenSession(docPath);

        // Modify properties in session
        _tool.Execute("set", sessionId: sessionId, title: "Modified Title");

        // Act - get properties from session
        var result = _tool.Execute("get", sessionId: sessionId);

        // Assert - should reflect the in-memory change
        Assert.Contains("Modified Title", result);
    }

    #endregion
}