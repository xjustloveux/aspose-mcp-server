using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Calendar;

/// <summary>
///     Information about a calendar appointment.
/// </summary>
public record AppointmentEmailInfo
{
    /// <summary>
    ///     Summary or title of the appointment.
    /// </summary>
    [JsonPropertyName("summary")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Summary { get; init; }

    /// <summary>
    ///     Detailed description of the appointment.
    /// </summary>
    [JsonPropertyName("description")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Description { get; init; }

    /// <summary>
    ///     Location of the appointment.
    /// </summary>
    [JsonPropertyName("location")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Location { get; init; }

    /// <summary>
    ///     Start date and time of the appointment in ISO 8601 format.
    /// </summary>
    [JsonPropertyName("startDate")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StartDate { get; init; }

    /// <summary>
    ///     End date and time of the appointment in ISO 8601 format.
    /// </summary>
    [JsonPropertyName("endDate")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EndDate { get; init; }

    /// <summary>
    ///     List of attendee email addresses.
    /// </summary>
    [JsonPropertyName("attendees")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<string>? Attendees { get; init; }

    /// <summary>
    ///     Whether the appointment has a recurrence pattern.
    /// </summary>
    [JsonPropertyName("isRecurring")]
    public required bool IsRecurring { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
