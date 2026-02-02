namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Collection definition for workflow integration tests.
///     Tests in this collection run serially to avoid resource contention
///     from concurrent heavy document processing operations.
/// </summary>
[CollectionDefinition("Workflow", DisableParallelization = true)]
public class WorkflowTestCollection;

/// <summary>
///     Collection definition for session integration tests.
///     Tests in this collection run serially to avoid race conditions
///     in session lifecycle management operations.
/// </summary>
[CollectionDefinition("Session Integration", DisableParallelization = true)]
public class SessionIntegrationTestCollection;
