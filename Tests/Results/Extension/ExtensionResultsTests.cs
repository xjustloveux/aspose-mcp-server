using System.Text.Json;
using AsposeMcpServer.Results.Extension;

namespace AsposeMcpServer.Tests.Results.Extension;

/// <summary>
///     Unit tests for Extension result DTOs.
/// </summary>
public class ExtensionResultsTests
{
    #region ExtensionResults Static Class Tests

    [Fact]
    public void ExtensionResults_AllTypes_ContainsExpectedTypes()
    {
        var types = ExtensionResults.AllTypes;

        Assert.Contains(typeof(ListExtensionsResult), types);
        Assert.Contains(typeof(ExtensionInfoDto), types);
        Assert.Contains(typeof(BindExtensionResult), types);
        Assert.Contains(typeof(UnbindExtensionResult), types);
        Assert.Contains(typeof(SetFormatResult), types);
        Assert.Contains(typeof(ExtensionStatusResult), types);
        Assert.Contains(typeof(ExtensionBindingsResult), types);
        Assert.Contains(typeof(BindingInfoDto), types);
    }

    [Fact]
    public void ExtensionResults_AllTypes_HasCorrectCount()
    {
        Assert.Equal(8, ExtensionResults.AllTypes.Length);
    }

    #endregion

    #region ListExtensionsResult Tests

    [Fact]
    public void ListExtensionsResult_SerializesToJson()
    {
        var result = new ListExtensionsResult
        {
            Success = true,
            Count = 2,
            Extensions = new List<ExtensionInfoDto>
            {
                new()
                {
                    Id = "ext1",
                    Name = "Extension 1",
                    Version = "1.0.0",
                    IsAvailable = true,
                    SupportedDocumentTypes = new List<string> { "word" },
                    InputFormats = new List<string> { "pdf" }
                },
                new()
                {
                    Id = "ext2",
                    Name = "Extension 2",
                    Version = "2.0.0",
                    IsAvailable = false,
                    UnavailableReason = "Not configured",
                    SupportedDocumentTypes = new List<string> { "excel" },
                    InputFormats = new List<string> { "html" }
                }
            }
        };

        var json = JsonSerializer.Serialize(result);

        Assert.Contains("\"success\":true", json);
        Assert.Contains("\"count\":2", json);
        Assert.Contains("\"extensions\":", json);
    }

    [Fact]
    public void ListExtensionsResult_DeserializesFromJson()
    {
        var json = """
                   {
                       "success": true,
                       "count": 1,
                       "extensions": [
                           {
                               "id": "test-ext",
                               "name": "Test Extension",
                               "version": "1.0.0",
                               "isAvailable": true,
                               "supportedDocumentTypes": ["word", "excel"],
                               "inputFormats": ["pdf"]
                           }
                       ]
                   }
                   """;

        var result = JsonSerializer.Deserialize<ListExtensionsResult>(json);

        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal(1, result.Count);
        Assert.Single(result.Extensions);
        Assert.Equal("test-ext", result.Extensions[0].Id);
    }

    #endregion

    #region ExtensionInfoDto Tests

    [Fact]
    public void ExtensionInfoDto_RequiredProperties_SetCorrectly()
    {
        var dto = new ExtensionInfoDto
        {
            Id = "pdf-viewer",
            Name = "PDF Viewer",
            Version = "1.2.3",
            Description = "Real-time PDF preview",
            IsAvailable = true,
            SupportedDocumentTypes = new List<string> { "word", "excel", "powerpoint" },
            InputFormats = new List<string> { "pdf", "png" },
            State = "Idle"
        };

        Assert.Equal("pdf-viewer", dto.Id);
        Assert.Equal("PDF Viewer", dto.Name);
        Assert.Equal("1.2.3", dto.Version);
        Assert.Equal("Real-time PDF preview", dto.Description);
        Assert.True(dto.IsAvailable);
        Assert.Equal(3, dto.SupportedDocumentTypes.Count);
        Assert.Equal(2, dto.InputFormats.Count);
        Assert.Equal("Idle", dto.State);
    }

    [Fact]
    public void ExtensionInfoDto_UnavailableExtension_HasReason()
    {
        var dto = new ExtensionInfoDto
        {
            Id = "broken-ext",
            Name = "Broken Extension",
            Version = "0.0.1",
            IsAvailable = false,
            UnavailableReason = "Executable not found",
            SupportedDocumentTypes = Array.Empty<string>(),
            InputFormats = Array.Empty<string>()
        };

        Assert.False(dto.IsAvailable);
        Assert.Equal("Executable not found", dto.UnavailableReason);
    }

    #endregion

    #region BindExtensionResult Tests

    [Fact]
    public void BindExtensionResult_Success_HasBinding()
    {
        var result = new BindExtensionResult
        {
            Success = true,
            Binding = new BindingInfoDto
            {
                SessionId = "sess_123",
                ExtensionId = "ext_456",
                OutputFormat = "pdf",
                CreatedAt = DateTime.UtcNow
            }
        };

        Assert.True(result.Success);
        Assert.NotNull(result.Binding);
        Assert.Null(result.Error);
    }

    [Fact]
    public void BindExtensionResult_Failure_HasError()
    {
        var result = new BindExtensionResult
        {
            Success = false,
            Error = "Session not found"
        };

        Assert.False(result.Success);
        Assert.Null(result.Binding);
        Assert.Equal("Session not found", result.Error);
    }

    [Fact]
    public void BindExtensionResult_SerializesToJson()
    {
        var result = new BindExtensionResult
        {
            Success = true,
            Binding = new BindingInfoDto
            {
                SessionId = "sess_abc",
                ExtensionId = "pdf-viewer",
                OutputFormat = "pdf",
                CreatedAt = DateTime.Parse("2024-01-15T10:30:00Z").ToUniversalTime()
            }
        };

        var json = JsonSerializer.Serialize(result);

        Assert.Contains("\"success\":true", json);
        Assert.Contains("\"binding\":", json);
        Assert.Contains("\"sessionId\":\"sess_abc\"", json);
    }

    #endregion

    #region UnbindExtensionResult Tests

    [Fact]
    public void UnbindExtensionResult_SingleUnbind_ReturnsCorrectCount()
    {
        var result = new UnbindExtensionResult
        {
            Success = true,
            SessionId = "sess_123",
            ExtensionId = "ext_456",
            UnboundCount = 1
        };

        Assert.True(result.Success);
        Assert.Equal("sess_123", result.SessionId);
        Assert.Equal("ext_456", result.ExtensionId);
        Assert.Equal(1, result.UnboundCount);
    }

    [Fact]
    public void UnbindExtensionResult_UnbindAll_ReturnsMultipleCount()
    {
        var result = new UnbindExtensionResult
        {
            Success = true,
            SessionId = "sess_123",
            ExtensionId = null,
            UnboundCount = 5
        };

        Assert.True(result.Success);
        Assert.Null(result.ExtensionId);
        Assert.Equal(5, result.UnboundCount);
    }

    #endregion

    #region SetFormatResult Tests

    [Fact]
    public void SetFormatResult_Success_HasAllProperties()
    {
        var result = new SetFormatResult
        {
            Success = true,
            SessionId = "sess_123",
            ExtensionId = "ext_456",
            NewFormat = "html"
        };

        Assert.True(result.Success);
        Assert.Null(result.Error);
        Assert.Equal("html", result.NewFormat);
    }

    [Fact]
    public void SetFormatResult_Failure_HasError()
    {
        var result = new SetFormatResult
        {
            Success = false,
            SessionId = "sess_123",
            ExtensionId = "ext_456",
            NewFormat = "xyz",
            Error = "Format not supported"
        };

        Assert.False(result.Success);
        Assert.Equal("Format not supported", result.Error);
    }

    #endregion

    #region ExtensionStatusResult Tests

    [Fact]
    public void ExtensionStatusResult_RunningExtension_HasCorrectState()
    {
        var result = new ExtensionStatusResult
        {
            Success = true,
            ExtensionId = "pdf-viewer",
            Name = "PDF Viewer",
            IsAvailable = true,
            State = "Idle",
            LastActivity = DateTime.UtcNow,
            RestartCount = 0,
            ActiveBindings = 3
        };

        Assert.True(result.Success);
        Assert.Equal("Idle", result.State);
        Assert.Equal(3, result.ActiveBindings);
    }

    [Fact]
    public void ExtensionStatusResult_NotFoundExtension_HasNotFoundState()
    {
        var result = new ExtensionStatusResult
        {
            Success = false,
            ExtensionId = "unknown",
            Name = "",
            State = "NotFound"
        };

        Assert.False(result.Success);
        Assert.Equal("NotFound", result.State);
    }

    #endregion

    #region ExtensionBindingsResult Tests

    [Fact]
    public void ExtensionBindingsResult_MultipleBindings_ListsAll()
    {
        var result = new ExtensionBindingsResult
        {
            Success = true,
            SessionId = "sess_123",
            Count = 2,
            Bindings = new List<BindingInfoDto>
            {
                new()
                {
                    SessionId = "sess_123",
                    ExtensionId = "ext_1",
                    OutputFormat = "pdf",
                    CreatedAt = DateTime.UtcNow
                },
                new()
                {
                    SessionId = "sess_123",
                    ExtensionId = "ext_2",
                    OutputFormat = "html",
                    CreatedAt = DateTime.UtcNow
                }
            }
        };

        Assert.True(result.Success);
        Assert.Equal(2, result.Count);
        Assert.Equal(2, result.Bindings.Count);
    }

    [Fact]
    public void ExtensionBindingsResult_NoBindings_ReturnsEmptyList()
    {
        var result = new ExtensionBindingsResult
        {
            Success = true,
            SessionId = "sess_123",
            Count = 0,
            Bindings = Array.Empty<BindingInfoDto>()
        };

        Assert.True(result.Success);
        Assert.Equal(0, result.Count);
        Assert.Empty(result.Bindings);
    }

    #endregion

    #region BindingInfoDto Tests

    [Fact]
    public void BindingInfoDto_AllProperties_SetCorrectly()
    {
        var createdAt = DateTime.Parse("2024-01-15T10:30:00Z").ToUniversalTime();
        var lastSentAt = DateTime.Parse("2024-01-15T10:35:00Z").ToUniversalTime();

        var dto = new BindingInfoDto
        {
            SessionId = "sess_abc123",
            ExtensionId = "pdf-viewer",
            OutputFormat = "pdf",
            CreatedAt = createdAt,
            LastSentAt = lastSentAt
        };

        Assert.Equal("sess_abc123", dto.SessionId);
        Assert.Equal("pdf-viewer", dto.ExtensionId);
        Assert.Equal("pdf", dto.OutputFormat);
        Assert.Equal(createdAt, dto.CreatedAt);
        Assert.Equal(lastSentAt, dto.LastSentAt);
    }

    [Fact]
    public void BindingInfoDto_SerializesCorrectly()
    {
        var dto = new BindingInfoDto
        {
            SessionId = "sess_test",
            ExtensionId = "ext_test",
            OutputFormat = "png",
            CreatedAt = DateTime.Parse("2024-01-01T00:00:00Z").ToUniversalTime()
        };

        var json = JsonSerializer.Serialize(dto);

        Assert.Contains("\"sessionId\":\"sess_test\"", json);
        Assert.Contains("\"extensionId\":\"ext_test\"", json);
        Assert.Contains("\"outputFormat\":\"png\"", json);
    }

    #endregion
}
