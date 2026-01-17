using System.Reflection;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Tests.Core.Handlers
{
    public class HandlerRegistryAutoDiscoveryTests
    {
        private static readonly Assembly TestAssembly = typeof(HandlerRegistryAutoDiscoveryTests).Assembly;

        #region CreateFromNamespace Tests

        [Fact]
        public void CreateFromNamespace_WithValidNamespace_DiscoversAllHandlers()
        {
            var registry = HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(
                "AsposeMcpServer.Tests.Core.Handlers.TestDiscoveryHandlers",
                TestAssembly);

            Assert.True(registry.Count >= 2);
            Assert.True(registry.HasHandler("discovery_add"));
            Assert.True(registry.HasHandler("discovery_delete"));
        }

        [Fact]
        public void CreateFromNamespace_ExcludesAttributedClasses()
        {
            var registry = HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(
                "AsposeMcpServer.Tests.Core.Handlers.TestDiscoveryHandlers",
                TestAssembly);

            Assert.True(registry.HasHandler("discovery_add"));
            Assert.True(registry.HasHandler("discovery_delete"));
            Assert.False(registry.HasHandler("excluded_operation"));
        }

        [Fact]
        public void CreateFromNamespace_WithNullNamespace_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() =>
                HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(null!, TestAssembly));
        }

        [Fact]
        public void CreateFromNamespace_WithEmptyNamespace_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() =>
                HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace("", TestAssembly));
        }

        [Fact]
        public void CreateFromNamespace_WithNoHandlers_ReturnsEmptyRegistry()
        {
            var registry = HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(
                "AsposeMcpServer.Tests.Core.Handlers.EmptyNamespace",
                TestAssembly);

            Assert.Equal(0, registry.Count);
        }

        [Fact]
        public void CreateFromNamespace_ExcludesAbstractClasses()
        {
            var registry = HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(
                "AsposeMcpServer.Tests.Core.Handlers.TestDiscoveryHandlers",
                TestAssembly);

            Assert.False(registry.HasHandler("abstract_operation"));
        }

        [Fact]
        public void CreateFromNamespace_ExcludesClassesWithoutParameterlessConstructor()
        {
            var registry = HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(
                "AsposeMcpServer.Tests.Core.Handlers.TestDiscoveryHandlers",
                TestAssembly);

            Assert.False(registry.HasHandler("no_parameterless_ctor"));
        }

        [Fact]
        public void CreateFromNamespace_HandlersAreExecutable()
        {
            var registry = HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(
                "AsposeMcpServer.Tests.Core.Handlers.TestDiscoveryHandlers",
                TestAssembly);

            var handler = registry.GetHandler("discovery_add");
            var doc = new TestDiscoveryDocument();
            var context = new OperationContext<TestDiscoveryDocument> { Document = doc };
            var parameters = new OperationParameters();

            var result = handler.Execute(context, parameters);

            Assert.Contains("discovery_add", result);
        }

        [Fact]
        public void CreateFromNamespace_WithDifferentContextTypes_OnlyDiscoversMatchingHandlers()
        {
            var registry = HandlerRegistry<TestDiscoveryDocument>.CreateFromNamespace(
                "AsposeMcpServer.Tests.Core.Handlers.TestDiscoveryHandlers",
                TestAssembly);

            Assert.False(registry.HasHandler("different_context_operation"));
        }

        #endregion
    }

    public class TestDiscoveryDocument
    {
        public string Content { get; set; } = "";
    }

    // ReSharper disable once ClassNeverInstantiated.Global - Used as generic type parameter for HandlerRegistry<T>
    public class DifferentContextDocument
    {
        public int Value { get; set; }
    }
}

namespace AsposeMcpServer.Tests.Core.Handlers.TestDiscoveryHandlers
{
    public class DiscoveryAddHandler : OperationHandlerBase<TestDiscoveryDocument>
    {
        public override string Operation => "discovery_add";

        public override string Execute(OperationContext<TestDiscoveryDocument> context, OperationParameters parameters)
        {
            return "Executed: discovery_add";
        }
    }

    public class DiscoveryDeleteHandler : OperationHandlerBase<TestDiscoveryDocument>
    {
        public override string Operation => "discovery_delete";

        public override string Execute(OperationContext<TestDiscoveryDocument> context, OperationParameters parameters)
        {
            return "Executed: discovery_delete";
        }
    }

    [ExcludeFromAutoDiscovery]
    public class ExcludedHandler : OperationHandlerBase<TestDiscoveryDocument>
    {
        public override string Operation => "excluded_operation";

        public override string Execute(OperationContext<TestDiscoveryDocument> context, OperationParameters parameters)
        {
            return "Should not be discovered";
        }
    }

    public abstract class AbstractHandler : OperationHandlerBase<TestDiscoveryDocument>
    {
        public override string Operation => "abstract_operation";
    }

    public class NoParameterlessCtorHandler : OperationHandlerBase<TestDiscoveryDocument>
    {
        private readonly string _requiredValue;

        public NoParameterlessCtorHandler(string requiredValue)
        {
            _requiredValue = requiredValue;
        }

        public override string Operation => "no_parameterless_ctor";

        public override string Execute(OperationContext<TestDiscoveryDocument> context, OperationParameters parameters)
        {
            return _requiredValue;
        }
    }

    public class DifferentContextHandler : OperationHandlerBase<DifferentContextDocument>
    {
        public override string Operation => "different_context_operation";

        public override string Execute(OperationContext<DifferentContextDocument> context,
            OperationParameters parameters)
        {
            return "Different context";
        }
    }
}
