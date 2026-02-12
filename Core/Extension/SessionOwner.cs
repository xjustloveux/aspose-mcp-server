using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Session owner information for group-based isolation.
/// </summary>
/// <remarks>
///     <para>Owner information is included based on the session isolation mode:</para>
///     <list type="bullet">
///         <item><c>IsolationMode=None</c>: Owner is null (no isolation)</item>
///         <item><c>IsolationMode=Group</c>: Only <see cref="GroupId" /> is populated</item>
///         <item><c>IsolationMode=User</c>: Both <see cref="GroupId" /> and <see cref="UserId" /> are populated</item>
///     </list>
///     <para>
///         Extensions can use this information to organize or classify documents
///         by tenant or user. The MCP server enforces isolation - extensions
///         only receive documents they're authorized to process.
///     </para>
/// </remarks>
public class SessionOwner
{
    /// <summary>
    ///     Group ID (tenant ID) for the session.
    ///     Has value when IsolationMode=Group or IsolationMode=User.
    /// </summary>
    [JsonPropertyName("groupId")]
    public string? GroupId { get; set; }

    /// <summary>
    ///     User ID for the session.
    ///     Has value only when IsolationMode=User.
    /// </summary>
    [JsonPropertyName("userId")]
    public string? UserId { get; set; }
}
