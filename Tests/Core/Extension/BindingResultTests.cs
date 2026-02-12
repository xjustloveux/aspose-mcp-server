using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Results.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for BindingResult class.
/// </summary>
public class BindingResultTests
{
    [Fact]
    public void BindingResult_Success_SetsPropertiesCorrectly()
    {
        var binding = new SessionBindingInfo
        {
            SessionId = "sess_123",
            ExtensionId = "ext_456",
            OutputFormat = "html"
        };

        var result = BindingResult.Success(binding);

        Assert.True(result.IsSuccess);
        Assert.NotNull(result.Binding);
        Assert.Equal("sess_123", result.Binding.SessionId);
        Assert.Equal("ext_456", result.Binding.ExtensionId);
        Assert.Equal(ExtensionErrorCode.None, result.ErrorCode);
        Assert.Null(result.Error);
    }

    [Fact]
    public void BindingResult_FailureWithErrorCode_SetsPropertiesCorrectly()
    {
        var result = BindingResult.Failure(ExtensionErrorCode.SessionNotFound, "Session not found");

        Assert.False(result.IsSuccess);
        Assert.Null(result.Binding);
        Assert.Equal(ExtensionErrorCode.SessionNotFound, result.ErrorCode);
        Assert.Equal("Session not found", result.Error);
    }

    [Fact]
    public void BindingResult_FailureWithoutErrorCode_UsesInternalError()
    {
        var result = BindingResult.Failure("Something went wrong");

        Assert.False(result.IsSuccess);
        Assert.Equal(ExtensionErrorCode.InternalError, result.ErrorCode);
        Assert.Equal("Something went wrong", result.Error);
    }

    [Fact]
    public void BindingResult_Success_BindingIsAccessible()
    {
        var binding = new SessionBindingInfo
        {
            SessionId = "test_session",
            ExtensionId = "test_extension",
            OutputFormat = "pdf",
            CreatedAt = DateTime.UtcNow
        };

        var result = BindingResult.Success(binding);

        Assert.Same(binding, result.Binding);
    }

    [Fact]
    public void BindingResult_Failure_ErrorMessagePreserved()
    {
        var errors = new[]
        {
            "Session not found: sess_abc",
            "Extension not found or unavailable: ext_xyz",
            "Format 'xyz' is not supported for document type 'word'"
        };

        foreach (var error in errors)
        {
            var result = BindingResult.Failure(error);

            Assert.Equal(error, result.Error);
        }
    }

    [Fact]
    public void BindingResult_Failure_EmptyErrorMessage()
    {
        var result = BindingResult.Failure(string.Empty);

        Assert.False(result.IsSuccess);
        Assert.Equal(string.Empty, result.Error);
    }

    [Fact]
    public void BindingResult_Success_WithMinimalBinding()
    {
        var binding = new SessionBindingInfo();

        var result = BindingResult.Success(binding);

        Assert.True(result.IsSuccess);
        Assert.NotNull(result.Binding);
        Assert.Equal(string.Empty, result.Binding.SessionId);
    }

    [Theory]
    [InlineData(ExtensionErrorCode.None)]
    [InlineData(ExtensionErrorCode.InvalidParameter)]
    [InlineData(ExtensionErrorCode.SessionNotFound)]
    [InlineData(ExtensionErrorCode.ExtensionNotFound)]
    [InlineData(ExtensionErrorCode.ExtensionUnavailable)]
    [InlineData(ExtensionErrorCode.BindingNotFound)]
    [InlineData(ExtensionErrorCode.FormatNotSupported)]
    [InlineData(ExtensionErrorCode.ConversionFailed)]
    [InlineData(ExtensionErrorCode.InternalError)]
    public void BindingResult_Failure_AllErrorCodes(ExtensionErrorCode errorCode)
    {
        var result = BindingResult.Failure(errorCode, "Test error");

        Assert.False(result.IsSuccess);
        Assert.Equal(errorCode, result.ErrorCode);
    }
}
