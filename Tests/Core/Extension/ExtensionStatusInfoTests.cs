using AsposeMcpServer.Core.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Tests for ExtensionStatusInfo class.
/// </summary>
public class ExtensionStatusInfoTests
{
    [Fact]
    public void Properties_CanBeSet()
    {
        var now = DateTime.UtcNow;
        var info = new ExtensionStatusInfo
        {
            Id = "test",
            Name = "Test Extension",
            IsAvailable = true,
            State = ExtensionState.Idle,
            LastActivity = now,
            RestartCount = 2,
            UnavailableReason = null
        };

        Assert.Equal("test", info.Id);
        Assert.Equal("Test Extension", info.Name);
        Assert.True(info.IsAvailable);
        Assert.Equal(ExtensionState.Idle, info.State);
        Assert.Equal(now, info.LastActivity);
        Assert.Equal(2, info.RestartCount);
        Assert.Null(info.UnavailableReason);
    }
}
