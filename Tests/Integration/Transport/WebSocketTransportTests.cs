using System.Net.WebSockets;
using AsposeMcpServer.Core.Transport;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Integration.Transport;

/// <summary>
///     Integration tests for WebSocket transport.
/// </summary>
[Trait("Category", "Integration")]
public class WebSocketTransportTests
{
    /// <summary>
    ///     Verifies that WebSocketConnectionHandler can be instantiated.
    /// </summary>
    [Fact]
    public void WebSocket_Handler_CanBeCreated()
    {
        var loggerFactory = Mock.Of<ILoggerFactory>();
        var handler = new WebSocketConnectionHandler("dotnet", "--all", loggerFactory);

        Assert.NotNull(handler);
    }

    /// <summary>
    ///     Verifies that WebSocketConnectionHandler handles null logger gracefully.
    /// </summary>
    [Fact]
    public void WebSocket_Handler_AcceptsNullLogger()
    {
        var handler = new WebSocketConnectionHandler("dotnet", "--all");

        Assert.NotNull(handler);
    }

    /// <summary>
    ///     Verifies that a closed WebSocket is handled gracefully.
    /// </summary>
    [Fact]
    public async Task WebSocket_ClosedConnection_HandledGracefully()
    {
        var handler = new WebSocketConnectionHandler("dotnet", "--all");
        var mockWebSocket = new Mock<WebSocket>();
        mockWebSocket.Setup(ws => ws.State).Returns(WebSocketState.Closed);
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));

        var exception =
            await Record.ExceptionAsync(() => handler.HandleConnectionAsync(mockWebSocket.Object, cts.Token));

        Assert.Null(exception);
    }

    /// <summary>
    ///     Verifies that cancellation is handled properly.
    /// </summary>
    [Fact]
    public async Task WebSocket_Cancellation_StopsProcessing()
    {
        var handler = new WebSocketConnectionHandler("dotnet", "--all");
        var mockWebSocket = new Mock<WebSocket>();
        mockWebSocket.Setup(ws => ws.State).Returns(WebSocketState.Open);
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        var task = handler.HandleConnectionAsync(mockWebSocket.Object, cts.Token);
        var completedInTime = await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(5))) == task;

        Assert.True(completedInTime);
    }

    /// <summary>
    ///     Verifies that group and user IDs are accepted.
    /// </summary>
    [Fact]
    public async Task WebSocket_WithIdentity_AcceptsGroupAndUser()
    {
        var handler = new WebSocketConnectionHandler("dotnet", "--all");
        var mockWebSocket = new Mock<WebSocket>();
        mockWebSocket.Setup(ws => ws.State).Returns(WebSocketState.Closed);
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(1));

        var exception = await Record.ExceptionAsync(() => handler.HandleConnectionAsync(
            mockWebSocket.Object,
            cts.Token,
            "test-group",
            "test-user"));

        Assert.Null(exception);
    }
}
