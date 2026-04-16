namespace AsposeMcpServer.Tests.Infrastructure.Ole;

/// <summary>
///     xUnit collection-fixture wiring for the OLE test suite so the 12-fixture matrix
///     is generated exactly once per test run (design §5.1). Test classes that need
///     fixtures attach via <c>[Collection(OleFixtureCollection.Name)]</c> and inject
///     <see cref="FixtureBuilder" /> through their constructor.
/// </summary>
[CollectionDefinition(Name)]
public sealed class OleFixtureCollection : ICollectionFixture<FixtureBuilder>
{
    /// <summary>Collection name referenced by test classes.</summary>
    public const string Name = "OLE Fixture Collection";
}
