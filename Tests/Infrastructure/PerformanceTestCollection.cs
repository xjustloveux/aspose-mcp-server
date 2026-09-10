namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Keeps wall-clock regression fixtures away from concurrent document processing. Their
///     correctness assertions still run normally; only the measured interval is isolated from
///     unrelated suite load.
/// </summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class PerformanceTestCollection
{
    /// <summary>The collection name used by performance regression fixtures.</summary>
    public const string Name = "Performance";
}
