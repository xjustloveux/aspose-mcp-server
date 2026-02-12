using System.Text.Json;
using AsposeMcpServer.Core.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Tests for <see cref="ExtensionInitializeResponse" />.
/// </summary>
public class ExtensionInitializeResponseTests
{
    /// <summary>
    ///     Tests that a valid JSON with required fields deserializes correctly.
    /// </summary>
    [Fact]
    public void Deserialize_ValidJson_ParsesCorrectly()
    {
        var json = """{"type":"initialize_response","name":"Test Extension","version":"1.0.0"}""";

        var response = JsonSerializer.Deserialize<ExtensionInitializeResponse>(json);

        Assert.NotNull(response);
        Assert.Equal("initialize_response", response.Type);
        Assert.Equal("Test Extension", response.Name);
        Assert.Equal("1.0.0", response.Version);
        Assert.Null(response.Title);
        Assert.Null(response.Description);
        Assert.Null(response.Author);
        Assert.Null(response.WebsiteUrl);
    }

    /// <summary>
    ///     Tests that a JSON with all optional fields deserializes correctly.
    /// </summary>
    [Fact]
    public void Deserialize_WithOptionalFields_ParsesAll()
    {
        var json = """
                   {
                       "type": "initialize_response",
                       "name": "Test Extension",
                       "version": "2.0.0",
                       "title": "測試擴充功能",
                       "description": "A test extension",
                       "author": "Test Author",
                       "websiteUrl": "https://example.com"
                   }
                   """;

        var response = JsonSerializer.Deserialize<ExtensionInitializeResponse>(json);

        Assert.NotNull(response);
        Assert.Equal("Test Extension", response.Name);
        Assert.Equal("2.0.0", response.Version);
        Assert.Equal("測試擴充功能", response.Title);
        Assert.Equal("A test extension", response.Description);
        Assert.Equal("Test Author", response.Author);
        Assert.Equal("https://example.com", response.WebsiteUrl);
    }

    /// <summary>
    ///     Tests that serialization produces correct JSON.
    /// </summary>
    [Fact]
    public void Serialize_ValidObject_ProducesCorrectJson()
    {
        var response = new ExtensionInitializeResponse
        {
            Name = "My Extension",
            Version = "1.2.3",
            Description = "Description"
        };

        var json = JsonSerializer.Serialize(response);

        Assert.Contains("\"type\":\"initialize_response\"", json);
        Assert.Contains("\"name\":\"My Extension\"", json);
        Assert.Contains("\"version\":\"1.2.3\"", json);
        Assert.Contains("\"description\":\"Description\"", json);
    }

    /// <summary>
    ///     Tests that default values are set correctly.
    /// </summary>
    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var response = new ExtensionInitializeResponse();

        Assert.Equal("initialize_response", response.Type);
        Assert.Equal(string.Empty, response.Name);
        Assert.Equal(string.Empty, response.Version);
        Assert.Null(response.Title);
        Assert.Null(response.Description);
        Assert.Null(response.Author);
        Assert.Null(response.WebsiteUrl);
    }

    /// <summary>
    ///     Tests that unknown fields are ignored during deserialization.
    /// </summary>
    [Fact]
    public void Deserialize_WithUnknownFields_IgnoresUnknown()
    {
        var json = """
                   {
                       "type": "initialize_response",
                       "name": "Test",
                       "version": "1.0.0",
                       "unknownField": "ignored",
                       "anotherUnknown": 123
                   }
                   """;

        var response = JsonSerializer.Deserialize<ExtensionInitializeResponse>(json);

        Assert.NotNull(response);
        Assert.Equal("Test", response.Name);
        Assert.Equal("1.0.0", response.Version);
    }
}
