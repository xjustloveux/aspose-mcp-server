using System.Text.Json;
using AsposeMcpServer.Core.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for ExtensionDefinition and related classes.
/// </summary>
public class ExtensionDefinitionTests
{
    #region ExtensionDefinition Default Values

    [Fact]
    public void ExtensionDefinition_DefaultValues_AreCorrect()
    {
        var definition = new ExtensionDefinition();

        Assert.Equal(string.Empty, definition.Id);
        Assert.Null(definition.RuntimeMetadata);
        Assert.Equal(string.Empty, definition.DisplayName); // Falls back to Id when no metadata
        Assert.Equal("unknown", definition.DisplayVersion);
        Assert.Null(definition.DisplayTitle);
        Assert.Null(definition.DisplayDescription);
        Assert.Null(definition.DisplayAuthor);
        Assert.Null(definition.DisplayWebsiteUrl);
        Assert.NotNull(definition.Command);
        Assert.Empty(definition.InputFormats);
        Assert.Empty(definition.SupportedDocumentTypes);
        Assert.Single(definition.TransportModes);
        Assert.Contains("file", definition.TransportModes);
        Assert.Null(definition.PreferredTransportMode);
        Assert.Equal("1.0", definition.ProtocolVersion);
        Assert.Null(definition.Capabilities);
        Assert.True(definition.IsAvailable);
        Assert.Null(definition.UnavailableReason);
    }

    [Fact]
    public void ExtensionDefinition_CanSetAllProperties()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test-ext",
            InputFormats = ["pdf", "html"],
            SupportedDocumentTypes = ["word", "excel"],
            TransportModes = ["mmap", "file"],
            PreferredTransportMode = "mmap",
            ProtocolVersion = "2.0",
            IsAvailable = false,
            UnavailableReason = "Test reason"
        };

        Assert.Equal("test-ext", definition.Id);
        Assert.Equal(2, definition.InputFormats.Count);
        Assert.Equal(2, definition.SupportedDocumentTypes.Count);
        Assert.Equal(2, definition.TransportModes.Count);
        Assert.Equal("mmap", definition.PreferredTransportMode);
        Assert.Equal("2.0", definition.ProtocolVersion);
        Assert.False(definition.IsAvailable);
        Assert.Equal("Test reason", definition.UnavailableReason);
    }

    [Fact]
    public void ExtensionDefinition_DisplayProperties_UseRuntimeMetadata()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test-ext",
            RuntimeMetadata = new ExtensionInitializeResponse
            {
                Name = "Test Extension",
                Version = "2.0.0",
                Title = "測試擴充",
                Description = "A test extension",
                Author = "Test Author",
                WebsiteUrl = "https://example.com"
            }
        };

        Assert.Equal("Test Extension", definition.DisplayName);
        Assert.Equal("2.0.0", definition.DisplayVersion);
        Assert.Equal("測試擴充", definition.DisplayTitle);
        Assert.Equal("A test extension", definition.DisplayDescription);
        Assert.Equal("Test Author", definition.DisplayAuthor);
        Assert.Equal("https://example.com", definition.DisplayWebsiteUrl);
    }

    [Fact]
    public void ExtensionDefinition_DisplayName_FallsBackToId_WhenNoMetadata()
    {
        var definition = new ExtensionDefinition { Id = "my-extension" };

        Assert.Equal("my-extension", definition.DisplayName);
        Assert.Equal("unknown", definition.DisplayVersion);
    }

    #endregion

    #region ExtensionCommand Tests

    [Fact]
    public void ExtensionCommand_DefaultValues_AreCorrect()
    {
        var command = new ExtensionCommand();

        Assert.Equal("executable", command.Type);
        Assert.Equal(string.Empty, command.Executable);
        Assert.Null(command.Arguments);
        Assert.Null(command.WorkingDirectory);
        Assert.Null(command.Environment);
    }

    [Fact]
    public void ExtensionCommand_CanSetAllProperties()
    {
        var command = new ExtensionCommand
        {
            Type = "python",
            Executable = "python",
            Arguments = "script.py --option value",
            WorkingDirectory = "/app/scripts",
            Environment = new Dictionary<string, string>
            {
                { "PYTHONPATH", "/app/lib" },
                { "DEBUG", "true" }
            }
        };

        Assert.Equal("python", command.Type);
        Assert.Equal("python", command.Executable);
        Assert.Equal("script.py --option value", command.Arguments);
        Assert.Equal("/app/scripts", command.WorkingDirectory);
        Assert.NotNull(command.Environment);
        Assert.Equal(2, command.Environment.Count);
        Assert.Equal("/app/lib", command.Environment["PYTHONPATH"]);
    }

    #endregion

    #region ExtensionCapabilities Tests

    [Fact]
    public void ExtensionCapabilities_DefaultValues_AreCorrect()
    {
        var capabilities = new ExtensionCapabilities();

        Assert.True(capabilities.SupportsHeartbeat);
        Assert.Null(capabilities.IdleTimeoutMinutes);
        Assert.Null(capabilities.FrameIntervalMs);
        Assert.Null(capabilities.MaxMissedHeartbeats);
        Assert.Null(capabilities.SnapshotTtlSeconds);
    }

    [Fact]
    public void ExtensionCapabilities_CanSetAllProperties()
    {
        var capabilities = new ExtensionCapabilities
        {
            SupportsHeartbeat = false,
            IdleTimeoutMinutes = 15,
            FrameIntervalMs = 200,
            MaxMissedHeartbeats = 5,
            SnapshotTtlSeconds = 60
        };

        Assert.False(capabilities.SupportsHeartbeat);
        Assert.Equal(15, capabilities.IdleTimeoutMinutes);
        Assert.Equal(200, capabilities.FrameIntervalMs);
        Assert.Equal(5, capabilities.MaxMissedHeartbeats);
        Assert.Equal(60, capabilities.SnapshotTtlSeconds);
    }

    #endregion

    #region ExtensionDefinition GetEffective Tests

    [Fact]
    public void GetEffectiveFrameIntervalMs_UsesExtensionValue_WhenWithinConstraints()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { FrameIntervalMs = 200 }
        };
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveFrameIntervalMs(config);

        Assert.Equal(200, effective);
    }

    [Fact]
    public void GetEffectiveFrameIntervalMs_UsesDefault_WhenNoExtensionValue()
    {
        var definition = new ExtensionDefinition();
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveFrameIntervalMs(config);

        Assert.Equal(config.FrameIntervalMs.Default, effective);
    }

    [Fact]
    public void GetEffectiveFrameIntervalMs_ClampsToFloor_WhenExtensionValueTooLow()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { FrameIntervalMs = 5 }
        };
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveFrameIntervalMs(config);

        Assert.Equal(config.FrameIntervalMs.Floor, effective);
    }

    [Fact]
    public void GetEffectiveFrameIntervalMs_ClampsToCeiling_WhenExtensionValueTooHigh()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { FrameIntervalMs = 10000 }
        };
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveFrameIntervalMs(config);

        Assert.Equal(config.FrameIntervalMs.Ceiling, effective);
    }

    [Fact]
    public void GetEffectiveSnapshotTtlSeconds_UsesExtensionValue_WhenWithinConstraints()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { SnapshotTtlSeconds = 60 }
        };
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveSnapshotTtlSeconds(config);

        Assert.Equal(60, effective);
    }

    [Fact]
    public void GetEffectiveSnapshotTtlSeconds_UsesDefault_WhenNoExtensionValue()
    {
        var definition = new ExtensionDefinition();
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveSnapshotTtlSeconds(config);

        Assert.Equal(config.SnapshotTtlSeconds.Default, effective);
    }

    [Fact]
    public void GetEffectiveMaxMissedHeartbeats_UsesExtensionValue_WhenWithinConstraints()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { MaxMissedHeartbeats = 5 }
        };
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveMaxMissedHeartbeats(config);

        Assert.Equal(5, effective);
    }

    [Fact]
    public void GetEffectiveMaxMissedHeartbeats_UsesDefault_WhenNoExtensionValue()
    {
        var definition = new ExtensionDefinition();
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveMaxMissedHeartbeats(config);

        Assert.Equal(config.MaxMissedHeartbeats.Default, effective);
    }

    [Fact]
    public void GetEffectiveIdleTimeoutMinutes_UsesExtensionValue_WhenWithinConstraints()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { IdleTimeoutMinutes = 15 }
        };
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveIdleTimeoutMinutes(config);

        Assert.Equal(15, effective);
    }

    [Fact]
    public void GetEffectiveIdleTimeoutMinutes_UsesDefault_WhenNoExtensionValue()
    {
        var definition = new ExtensionDefinition();
        var config = new ExtensionConfig();

        var effective = definition.GetEffectiveIdleTimeoutMinutes(config);

        Assert.Equal(config.IdleTimeoutMinutes.Default, effective);
    }

    [Fact]
    public void GetEffectiveIdleTimeoutMinutes_AllowsZero_WhenSpecialValueAllowed()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { IdleTimeoutMinutes = 0 }
        };
        var config = new ExtensionConfig
        {
            IdleTimeoutMinutes = { SpecialAllowed = true }
        };

        var effective = definition.GetEffectiveIdleTimeoutMinutes(config);

        Assert.Equal(0, effective);
    }

    [Fact]
    public void GetEffectiveIdleTimeoutMinutes_ReturnsDefault_WhenSpecialValueNotAllowed()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { IdleTimeoutMinutes = 0 }
        };
        var config = new ExtensionConfig
        {
            IdleTimeoutMinutes = { SpecialAllowed = false }
        };

        var effective = definition.GetEffectiveIdleTimeoutMinutes(config);

        Assert.Equal(config.IdleTimeoutMinutes.Default, effective);
    }

    [Fact]
    public void GetEffective_WorksWithNullCapabilities()
    {
        var definition = new ExtensionDefinition { Capabilities = null };
        var config = new ExtensionConfig();

        Assert.Equal(config.FrameIntervalMs.Default, definition.GetEffectiveFrameIntervalMs(config));
        Assert.Equal(config.SnapshotTtlSeconds.Default, definition.GetEffectiveSnapshotTtlSeconds(config));
        Assert.Equal(config.MaxMissedHeartbeats.Default, definition.GetEffectiveMaxMissedHeartbeats(config));
        Assert.Equal(config.IdleTimeoutMinutes.Default, definition.GetEffectiveIdleTimeoutMinutes(config));
    }

    #endregion

    #region ValidateCapabilityConstraints Tests

    [Fact]
    public void ValidateCapabilityConstraints_AllValuesValid_ReturnsEmpty()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities
            {
                FrameIntervalMs = 200,
                SnapshotTtlSeconds = 60,
                MaxMissedHeartbeats = 5
            }
        };
        var config = new ExtensionConfig();

        var warnings = definition.ValidateCapabilityConstraints(config);

        Assert.Empty(warnings);
    }

    [Fact]
    public void ValidateCapabilityConstraints_NullCapabilities_ReturnsEmpty()
    {
        var definition = new ExtensionDefinition { Capabilities = null };
        var config = new ExtensionConfig();

        var warnings = definition.ValidateCapabilityConstraints(config);

        Assert.Empty(warnings);
    }

    [Fact]
    public void ValidateCapabilityConstraints_ValueBelowFloor_ReturnsWarning()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { FrameIntervalMs = 1 }
        };
        var config = new ExtensionConfig();

        var warnings = definition.ValidateCapabilityConstraints(config);

        Assert.Single(warnings);
        Assert.Contains("FrameIntervalMs=1", warnings[0]);
        Assert.Contains("below minimum", warnings[0]);
    }

    [Fact]
    public void ValidateCapabilityConstraints_ValueAboveCeiling_ReturnsWarning()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { SnapshotTtlSeconds = 1000 }
        };
        var config = new ExtensionConfig();

        var warnings = definition.ValidateCapabilityConstraints(config);

        Assert.Single(warnings);
        Assert.Contains("SnapshotTtlSeconds=1000", warnings[0]);
        Assert.Contains("above maximum", warnings[0]);
    }

    [Fact]
    public void ValidateCapabilityConstraints_MultipleViolations_ReturnsAllWarnings()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities
            {
                FrameIntervalMs = 1,
                SnapshotTtlSeconds = 1000,
                MaxMissedHeartbeats = 100
            }
        };
        var config = new ExtensionConfig();

        var warnings = definition.ValidateCapabilityConstraints(config);

        Assert.Equal(3, warnings.Count);
    }

    [Fact]
    public void ValidateCapabilityConstraints_SpecialValueNotAllowed_ReturnsWarning()
    {
        var definition = new ExtensionDefinition
        {
            Capabilities = new ExtensionCapabilities { IdleTimeoutMinutes = 0 }
        };
        var config = new ExtensionConfig
        {
            IdleTimeoutMinutes = { SpecialAllowed = false }
        };

        var warnings = definition.ValidateCapabilityConstraints(config);

        Assert.Single(warnings);
        Assert.Contains("special value", warnings[0]);
        Assert.Contains("not allowed", warnings[0]);
    }

    #endregion

    #region ExtensionsConfigFile Tests

    [Fact]
    public void ExtensionsConfigFile_DefaultValues_AreCorrect()
    {
        var configFile = new ExtensionsConfigFile();

        Assert.NotNull(configFile.Extensions);
        Assert.Empty(configFile.Extensions);
    }

    [Fact]
    public void ExtensionsConfigFile_CanAddExtensions()
    {
        var configFile = new ExtensionsConfigFile
        {
            Extensions = new Dictionary<string, ExtensionDefinition>
            {
                ["ext1"] = new(),
                ["ext2"] = new()
            }
        };

        Assert.Equal(2, configFile.Extensions.Count);
        Assert.True(configFile.Extensions.ContainsKey("ext1"));
        Assert.True(configFile.Extensions.ContainsKey("ext2"));
    }

    #endregion

    #region JSON Serialization Tests

    [Fact]
    public void ExtensionDefinition_SerializesToJson_WithCorrectPropertyNames()
    {
        var definition = new ExtensionDefinition
        {
            InputFormats = ["pdf"],
            SupportedDocumentTypes = ["word"]
        };

        var json = JsonSerializer.Serialize(definition);

        Assert.Contains("\"inputFormats\":", json);
        Assert.Contains("\"supportedDocumentTypes\":", json);
        Assert.Contains("\"transportModes\":", json);
        Assert.Contains("\"protocolVersion\":", json);

        Assert.DoesNotContain("\"id\":", json);
        Assert.DoesNotContain("\"runtimeMetadata\":", json);
    }

    [Fact]
    public void ExtensionDefinition_DeserializesFromJson()
    {
        var json = """
                   {
                       "inputFormats": ["pdf", "png"],
                       "supportedDocumentTypes": ["word", "excel", "powerpoint"],
                       "transportModes": ["mmap", "file"],
                       "preferredTransportMode": "mmap",
                       "protocolVersion": "1.5",
                       "command": {
                           "type": "node",
                           "executable": "node",
                           "arguments": "viewer.js"
                       }
                   }
                   """;

        var definition = JsonSerializer.Deserialize<ExtensionDefinition>(json);

        Assert.NotNull(definition);
        Assert.Equal(2, definition.InputFormats.Count);
        Assert.Contains("pdf", definition.InputFormats);
        Assert.Equal(3, definition.SupportedDocumentTypes.Count);
        Assert.Equal(2, definition.TransportModes.Count);
        Assert.Equal("mmap", definition.PreferredTransportMode);
        Assert.Equal("1.5", definition.ProtocolVersion);
        Assert.Equal("node", definition.Command.Type);
        Assert.Equal("node", definition.Command.Executable);
        Assert.Equal("viewer.js", definition.Command.Arguments);
    }

    [Fact]
    public void ExtensionsConfigFile_DeserializesFromJson()
    {
        var json = """
                   {
                       "extensions": {
                           "ext1": {
                               "inputFormats": ["pdf"]
                           },
                           "ext2": {
                               "inputFormats": ["html"]
                           }
                       }
                   }
                   """;

        var configFile = JsonSerializer.Deserialize<ExtensionsConfigFile>(json);

        Assert.NotNull(configFile);
        Assert.Equal(2, configFile.Extensions.Count);
        Assert.True(configFile.Extensions.ContainsKey("ext1"));
        Assert.True(configFile.Extensions.ContainsKey("ext2"));
        Assert.Contains("pdf", configFile.Extensions["ext1"].InputFormats);
        Assert.Contains("html", configFile.Extensions["ext2"].InputFormats);
    }

    [Fact]
    public void ExtensionCommand_DeserializesFromJson_WithEnvironment()
    {
        var json = """
                   {
                       "type": "python",
                       "executable": "python3",
                       "arguments": "-u script.py",
                       "workingDirectory": "/app",
                       "environment": {
                           "PYTHONPATH": "/app/lib",
                           "LOG_LEVEL": "debug"
                       }
                   }
                   """;

        var command = JsonSerializer.Deserialize<ExtensionCommand>(json);

        Assert.NotNull(command);
        Assert.Equal("python", command.Type);
        Assert.Equal("python3", command.Executable);
        Assert.Equal("-u script.py", command.Arguments);
        Assert.Equal("/app", command.WorkingDirectory);
        Assert.NotNull(command.Environment);
        Assert.Equal(2, command.Environment.Count);
        Assert.Equal("/app/lib", command.Environment["PYTHONPATH"]);
        Assert.Equal("debug", command.Environment["LOG_LEVEL"]);
    }

    [Fact]
    public void ExtensionDefinition_RuntimeProperties_NotSerialized()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = false,
            UnavailableReason = "Test reason",
            RuntimeMetadata = new ExtensionInitializeResponse
            {
                Name = "Test",
                Version = "1.0.0"
            }
        };

        var json = JsonSerializer.Serialize(definition);

        Assert.DoesNotContain("IsAvailable", json);
        Assert.DoesNotContain("isAvailable", json);
        Assert.DoesNotContain("UnavailableReason", json);
        Assert.DoesNotContain("unavailableReason", json);
        Assert.DoesNotContain("runtimeMetadata", json);
        Assert.DoesNotContain("RuntimeMetadata", json);
        Assert.DoesNotContain("\"id\"", json);
    }

    [Fact]
    public void ExtensionCapabilities_SerializesCorrectly()
    {
        var capabilities = new ExtensionCapabilities
        {
            SupportsHeartbeat = true,
            FrameIntervalMs = 200
        };

        var json = JsonSerializer.Serialize(capabilities);
        var deserialized = JsonSerializer.Deserialize<ExtensionCapabilities>(json);

        Assert.NotNull(deserialized);
        Assert.True(deserialized.SupportsHeartbeat);
        Assert.Equal(200, deserialized.FrameIntervalMs);
    }

    [Fact]
    public void ExtensionCapabilities_SerializesOverridableSettings()
    {
        var capabilities = new ExtensionCapabilities
        {
            IdleTimeoutMinutes = 15,
            FrameIntervalMs = 200,
            MaxMissedHeartbeats = 5,
            SnapshotTtlSeconds = 60
        };

        var json = JsonSerializer.Serialize(capabilities);

        Assert.Contains("\"idleTimeoutMinutes\":15", json);
        Assert.Contains("\"frameIntervalMs\":200", json);
        Assert.Contains("\"maxMissedHeartbeats\":5", json);
        Assert.Contains("\"snapshotTtlSeconds\":60", json);
    }

    [Fact]
    public void ExtensionCapabilities_DeserializesOverridableSettings()
    {
        var json = """
                   {
                       "supportsHeartbeat": true,
                       "idleTimeoutMinutes": 15,
                       "frameIntervalMs": 200,
                       "maxMissedHeartbeats": 5,
                       "snapshotTtlSeconds": 60
                   }
                   """;

        var capabilities = JsonSerializer.Deserialize<ExtensionCapabilities>(json);

        Assert.NotNull(capabilities);
        Assert.Equal(15, capabilities.IdleTimeoutMinutes);
        Assert.Equal(200, capabilities.FrameIntervalMs);
        Assert.Equal(5, capabilities.MaxMissedHeartbeats);
        Assert.Equal(60, capabilities.SnapshotTtlSeconds);
    }

    [Fact]
    public void ExtensionCapabilities_NullOverridableSettings_NotSerialized()
    {
        var capabilities = new ExtensionCapabilities();

        var json = JsonSerializer.Serialize(capabilities);

        Assert.DoesNotContain("idleTimeoutMinutes", json);
        Assert.DoesNotContain("frameIntervalMs", json);
        Assert.DoesNotContain("maxMissedHeartbeats", json);
        Assert.DoesNotContain("snapshotTtlSeconds", json);
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void ExtensionDefinition_EmptyInputFormats_HandledCorrectly()
    {
        var json = """
                   {
                       "inputFormats": []
                   }
                   """;

        var definition = JsonSerializer.Deserialize<ExtensionDefinition>(json);

        Assert.NotNull(definition);
        Assert.Empty(definition.InputFormats);
    }

    [Fact]
    public void ExtensionDefinition_MissingOptionalFields_UsesDefaults()
    {
        var json = """
                   {
                       "command": {
                           "type": "node",
                           "executable": "index.js"
                       }
                   }
                   """;

        var definition = JsonSerializer.Deserialize<ExtensionDefinition>(json);

        Assert.NotNull(definition);
        Assert.Equal("1.0", definition.ProtocolVersion);
        Assert.Single(definition.TransportModes);
        Assert.Contains("file", definition.TransportModes);
    }

    #endregion
}
