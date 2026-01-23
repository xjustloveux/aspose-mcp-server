using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Tests.Core.Handlers;

/// <summary>
///     Unit tests for HandlerRegistry
/// </summary>
public class HandlerRegistryTests
{
    [Fact]
    public void Register_AddsHandler()
    {
        var registry = new HandlerRegistry<TestDocument>();
        var handler = new TestAddHandler();

        registry.Register(handler);

        Assert.Equal(1, registry.Count);
        Assert.True(registry.HasHandler("add"));
    }

    [Fact]
    public void Register_MultipleHandlers_AllRegistered()
    {
        var registry = new HandlerRegistry<TestDocument>();

        registry.Register(new TestAddHandler());
        registry.Register(new TestDeleteHandler());
        registry.Register(new TestSearchHandler());

        Assert.Equal(3, registry.Count);
        Assert.True(registry.HasHandler("add"));
        Assert.True(registry.HasHandler("delete"));
        Assert.True(registry.HasHandler("search"));
    }

    [Fact]
    public void Register_DuplicateOperation_ThrowsInvalidOperationException()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());

        var ex = Assert.Throws<InvalidOperationException>(() => registry.Register(new TestAddHandler()));

        Assert.Contains("add", ex.Message);
        Assert.Contains("already registered", ex.Message);
    }

    [Fact]
    public void GetHandler_WithRegisteredOperation_ReturnsHandler()
    {
        var registry = new HandlerRegistry<TestDocument>();
        var expectedHandler = new TestAddHandler();
        registry.Register(expectedHandler);

        var handler = registry.GetHandler("add");

        Assert.Same(expectedHandler, handler);
    }

    [Fact]
    public void GetHandler_IsCaseInsensitive()
    {
        var registry = new HandlerRegistry<TestDocument>();
        var expectedHandler = new TestAddHandler();
        registry.Register(expectedHandler);

        Assert.Same(expectedHandler, registry.GetHandler("add"));
        Assert.Same(expectedHandler, registry.GetHandler("ADD"));
        Assert.Same(expectedHandler, registry.GetHandler("Add"));
    }

    [Fact]
    public void GetHandler_WithUnknownOperation_ThrowsArgumentException()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());
        registry.Register(new TestDeleteHandler());

        var ex = Assert.Throws<ArgumentException>(() => registry.GetHandler("unknown"));

        Assert.Contains("Unknown operation: unknown", ex.Message);
        Assert.Contains("Available operations:", ex.Message);
        Assert.Contains("add", ex.Message);
        Assert.Contains("delete", ex.Message);
    }

    [Fact]
    public void TryGetHandler_WithRegisteredOperation_ReturnsTrueAndHandler()
    {
        var registry = new HandlerRegistry<TestDocument>();
        var expectedHandler = new TestAddHandler();
        registry.Register(expectedHandler);

        var found = registry.TryGetHandler("add", out var handler);

        Assert.True(found);
        Assert.Same(expectedHandler, handler);
    }

    [Fact]
    public void TryGetHandler_WithUnknownOperation_ReturnsFalseAndNull()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());

        var found = registry.TryGetHandler("unknown", out var handler);

        Assert.False(found);
        Assert.Null(handler);
    }

    [Fact]
    public void TryGetHandler_IsCaseInsensitive()
    {
        var registry = new HandlerRegistry<TestDocument>();
        var expectedHandler = new TestDeleteHandler();
        registry.Register(expectedHandler);

        Assert.True(registry.TryGetHandler("DELETE", out var handler1));
        Assert.Same(expectedHandler, handler1);

        Assert.True(registry.TryGetHandler("Delete", out var handler2));
        Assert.Same(expectedHandler, handler2);
    }

    [Fact]
    public void GetOperations_ReturnsAllRegisteredOperations()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());
        registry.Register(new TestDeleteHandler());
        registry.Register(new TestSearchHandler());

        var operations = registry.GetOperations().ToList();

        Assert.Equal(3, operations.Count);
        Assert.Contains("add", operations);
        Assert.Contains("delete", operations);
        Assert.Contains("search", operations);
    }

    [Fact]
    public void GetOperations_WithEmptyRegistry_ReturnsEmpty()
    {
        var registry = new HandlerRegistry<TestDocument>();

        var operations = registry.GetOperations();

        Assert.Empty(operations);
    }

    [Fact]
    public void Count_WithEmptyRegistry_ReturnsZero()
    {
        var registry = new HandlerRegistry<TestDocument>();

        Assert.Equal(0, registry.Count);
    }

    [Fact]
    public void HasHandler_WithRegisteredOperation_ReturnsTrue()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());

        Assert.True(registry.HasHandler("add"));
    }

    [Fact]
    public void HasHandler_WithUnknownOperation_ReturnsFalse()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());

        Assert.False(registry.HasHandler("unknown"));
    }

    [Fact]
    public void HasHandler_IsCaseInsensitive()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());

        Assert.True(registry.HasHandler("ADD"));
        Assert.True(registry.HasHandler("Add"));
        Assert.True(registry.HasHandler("add"));
    }

    [Fact]
    public void Handler_Execute_WorksCorrectly()
    {
        var registry = new HandlerRegistry<TestDocument>();
        registry.Register(new TestAddHandler());

        var handler = registry.GetHandler("add");
        var doc = new TestDocument();
        var context = new OperationContext<TestDocument> { Document = doc };
        var parameters = new OperationParameters();
        parameters.Set("text", "hello");

        var result = handler.Execute(context, parameters);

        Assert.Equal("Added: hello", result);
        Assert.True(context.IsModified);
        Assert.Equal("hello", doc.Content);
    }

    private class TestDocument
    {
        public string Content { get; set; } = "";
    }

    private class TestAddHandler : OperationHandlerBase<TestDocument>
    {
        public override string Operation => "add";

        public override object Execute(OperationContext<TestDocument> context, OperationParameters parameters)
        {
            var text = parameters.GetRequired<string>("text");
            context.Document.Content = text;
            MarkModified(context);
            return $"Added: {text}";
        }
    }

    private class TestDeleteHandler : OperationHandlerBase<TestDocument>
    {
        public override string Operation => "delete";

        public override object Execute(OperationContext<TestDocument> context, OperationParameters parameters)
        {
            context.Document.Content = "";
            MarkModified(context);
            return "Deleted";
        }
    }

    private class TestSearchHandler : OperationHandlerBase<TestDocument>
    {
        public override string Operation => "search";

        public override object Execute(OperationContext<TestDocument> context, OperationParameters parameters)
        {
            var searchText = parameters.GetRequired<string>("searchText");
            var found = context.Document.Content.Contains(searchText);
            return new { found, searchText };
        }
    }
}
