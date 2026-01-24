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
        // Arrange
        var loggerFactory = Mock.Of<ILoggerFactory>();

        // Act
        var handler = new WebSocketConnectionHandler("dotnet", "--all", loggerFactory);

        // Assert
        Assert.NotNull(handler);
    }

    /// <summary>
    ///     Verifies that WebSocketConnectionHandler handles null logger gracefully.
    /// </summary>
    [Fact]
    public void WebSocket_Handler_AcceptsNullLogger()
    {
        // Act
        var handler = new WebSocketConnectionHandler("dotnet", "--all");

        // Assert
        Assert.NotNull(handler);
    }

    /// <summary>
    ///     Verifies that a closed WebSocket is handled gracefully.
    /// </summary>
    [Fact]
    public async Task WebSocket_ClosedConnection_HandledGracefully()
    {
        // Arrange
        var handler = new WebSocketConnectionHandler("dotnet", "--all");
        var mockWebSocket = new Mock<WebSocket>();
        mockWebSocket.Setup(ws => ws.State).Returns(WebSocketState.Closed);

        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));

        // Act & Assert - should not throw
        await handler.HandleConnectionAsync(mockWebSocket.Object, cts.Token);
    }

    /// <summary>
    ///     Verifies that cancellation is handled properly.
    /// </summary>
    [Fact]
    public async Task WebSocket_Cancellation_StopsProcessing()
    {
        // Arrange
        var handler = new WebSocketConnectionHandler("dotnet", "--all");
        var mockWebSocket = new Mock<WebSocket>();
        mockWebSocket.Setup(ws => ws.State).Returns(WebSocketState.Open);

        using var cts = new CancellationTokenSource();

        // Cancel immediately
        await cts.CancelAsync();

        // Act & Assert - should complete quickly due to cancellation
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
        // Arrange
        var handler = new WebSocketConnectionHandler("dotnet", "--all");
        var mockWebSocket = new Mock<WebSocket>();
        mockWebSocket.Setup(ws => ws.State).Returns(WebSocketState.Closed);

        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(1));

        // Act & Assert - should accept identity parameters
        await handler.HandleConnectionAsync(
            mockWebSocket.Object,
            cts.Token,
            "test-group",
            "test-user");
    }
}
