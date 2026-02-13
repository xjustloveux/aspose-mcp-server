using System.IO.Hashing;
using AsposeMcpServer.Core.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

public class ExtensionMetadataTests
{
    [Fact]
    public void VerifyChecksum_ValidData_ReturnsTrue()
    {
        var data = "Hello, World!"u8.ToArray();
        var metadata = new ExtensionMetadata
        {
            DataSize = data.Length,
            Checksum = Crc32.HashToUInt32(data)
        };

        var result = metadata.VerifyChecksum(data);

        Assert.True(result);
    }

    [Fact]
    public void VerifyChecksum_InvalidChecksum_ReturnsFalse()
    {
        var data = "Hello, World!"u8.ToArray();
        var metadata = new ExtensionMetadata
        {
            DataSize = data.Length,
            Checksum = 12345
        };

        var result = metadata.VerifyChecksum(data);

        Assert.False(result);
    }

    [Fact]
    public void VerifyChecksum_EmptyData_ZeroChecksum_ReturnsTrue()
    {
        var data = Array.Empty<byte>();
        var metadata = new ExtensionMetadata
        {
            DataSize = 0,
            Checksum = 0
        };

        var result = metadata.VerifyChecksum(data);

        Assert.True(result);
    }

    [Fact]
    public void VerifyChecksum_NullData_ZeroChecksum_ReturnsTrue()
    {
        var metadata = new ExtensionMetadata
        {
            DataSize = 0,
            Checksum = 0
        };

        var result = metadata.VerifyChecksum(null!);

        Assert.True(result);
    }

    [Fact]
    public void VerifyData_ValidData_ReturnsValid()
    {
        var data = "Test data for verification"u8.ToArray();
        var metadata = new ExtensionMetadata
        {
            DataSize = data.Length,
            Checksum = Crc32.HashToUInt32(data)
        };

        var result = metadata.VerifyData(data);

        Assert.Equal(DataVerificationResult.Valid, result);
    }

    [Fact]
    public void VerifyData_NullData_ReturnsNullData()
    {
        var metadata = new ExtensionMetadata
        {
            DataSize = 10,
            Checksum = 12345
        };

        var result = metadata.VerifyData(null!);

        Assert.Equal(DataVerificationResult.NullData, result);
    }

    [Fact]
    public void VerifyData_WrongSize_ReturnsSizeMismatch()
    {
        var data = "Short"u8.ToArray();
        var metadata = new ExtensionMetadata
        {
            DataSize = 100,
            Checksum = Crc32.HashToUInt32(data)
        };

        var result = metadata.VerifyData(data);

        Assert.Equal(DataVerificationResult.SizeMismatch, result);
    }

    [Fact]
    public void VerifyData_WrongChecksum_ReturnsChecksumMismatch()
    {
        var data = "Test data"u8.ToArray();
        var metadata = new ExtensionMetadata
        {
            DataSize = data.Length,
            Checksum = 12345
        };

        var result = metadata.VerifyData(data);

        Assert.Equal(DataVerificationResult.ChecksumMismatch, result);
    }

    [Fact]
    public void VerifyData_CorruptedData_ReturnsChecksumMismatch()
    {
        var originalData = "Original data"u8.ToArray();
        var corruptedData = "Corrupted d"u8.ToArray();
        var metadata = new ExtensionMetadata
        {
            DataSize = originalData.Length,
            Checksum = Crc32.HashToUInt32(originalData)
        };

        var result = metadata.VerifyData(corruptedData);

        Assert.Equal(DataVerificationResult.SizeMismatch, result);
    }

    [Fact]
    public void VerifyData_SameSizeDifferentContent_ReturnsChecksumMismatch()
    {
        var originalData = "ABCD"u8.ToArray();
        var differentData = "EFGH"u8.ToArray();
        var metadata = new ExtensionMetadata
        {
            DataSize = originalData.Length,
            Checksum = Crc32.HashToUInt32(originalData)
        };

        var result = metadata.VerifyData(differentData);

        Assert.Equal(DataVerificationResult.ChecksumMismatch, result);
    }

    [Fact]
    public void VerifyChecksum_NullData_NonZeroChecksum_ReturnsFalse()
    {
        var metadata = new ExtensionMetadata
        {
            DataSize = 0,
            Checksum = 12345
        };

        var result = metadata.VerifyChecksum(null);

        Assert.False(result);
    }

    [Fact]
    public void VerifyChecksum_EmptyData_NonZeroChecksum_ReturnsFalse()
    {
        var data = Array.Empty<byte>();
        var metadata = new ExtensionMetadata
        {
            DataSize = 0,
            Checksum = 12345
        };

        var result = metadata.VerifyChecksum(data);

        Assert.False(result);
    }

    [Fact]
    public void CreateCommand_ReturnsValidCommandMetadata()
    {
        var sessionId = "test-session-123";
        var commandType = "highlight";
        var payload = new Dictionary<string, object> { { "color", "yellow" } };

        var result = ExtensionMetadata.CreateCommand(sessionId, commandType, payload);

        Assert.Equal("command", result.Type);
        Assert.Equal(sessionId, result.SessionId);
        Assert.Equal(commandType, result.CommandType);
        Assert.Equal(payload, result.CommandPayload);
        Assert.NotNull(result.CommandId);
        Assert.NotEmpty(result.CommandId);
    }

    [Fact]
    public void CreateCommand_WithoutPayload_ReturnsValidCommandMetadata()
    {
        var sessionId = "test-session-456";
        var commandType = "navigate";

        var result = ExtensionMetadata.CreateCommand(sessionId, commandType);

        Assert.Equal("command", result.Type);
        Assert.Equal(sessionId, result.SessionId);
        Assert.Equal(commandType, result.CommandType);
        Assert.Null(result.CommandPayload);
        Assert.NotNull(result.CommandId);
    }

    [Fact]
    public void CreateCommand_GeneratesUniqueCommandIds()
    {
        var command1 = ExtensionMetadata.CreateCommand("session1", "type1");
        var command2 = ExtensionMetadata.CreateCommand("session2", "type2");

        Assert.NotEqual(command1.CommandId, command2.CommandId);
    }
}
