namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Collection for fixtures that replace a <em>static</em> seam on production code.
///     <para>
///         <c>BoundedFileBatch.RecordDebt</c> is one field for the whole process. Two test classes
///         that each install their own collector run in parallel by default, so one class's batch
///         reports into the other's list: the collector sees debts it never caused and the
///         control that asserts "an ordinary publish queues nothing" fails. Both classes passed in
///         isolation and failed together, which is the signature of shared static state rather
///         than of either test being wrong.
///     </para>
///     <para>
///         This was reachable as soon as a second class used the seam. Anything that assigns a
///         static seam belongs here; a per-instance seam does not need it.
///     </para>
/// </summary>
[CollectionDefinition("SerialStaticSeams", DisableParallelization = true)]
public class SerialStaticSeamTestCollection;
