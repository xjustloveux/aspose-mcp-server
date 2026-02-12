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
}
