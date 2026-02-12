using AsposeMcpServer.Core.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Tests for ExtensionMessageEventArgs class.
/// </summary>
public class ExtensionMessageEventArgsTests
{
    [Fact]
    public void Constructor_SetsProperties()
    {
        var args = new ExtensionMessageEventArgs("test_type", "{\"data\":1}");

        Assert.Equal("test_type", args.MessageType);
        Assert.Equal("{\"data\":1}", args.RawJson);
    }
}
