using Aspose.Email.Calendar;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Handlers.Email;

/// <summary>
///     Covers LOW-09. Dates arrive as ISO 8601 strings and were parsed with
///     <c>DateTimeStyles.None</c>, which converts a trailing <c>Z</c> to the server's local time
///     and marks the result <c>Local</c>. Reading it back on the same machine hides the loss; the
///     written file is what shows it, because a local-kind time is serialised without the UTC
///     marker and means a different instant on a reader in another timezone.
/// </summary>
public class CalendarUtcRoundTripTests : TestBase
{
    [Fact]
    public void CreateAppointment_WithUtcInstant_ShouldWriteAUtcTimestamp()
    {
        var tool = new EmailCalendarTool();
        var outputPath = CreateTestFilePath("utc.ics");

        tool.Execute("create", outputPath: outputPath,
            summary: "Review",
            startDate: "2026-09-04T10:00:00Z",
            endDate: "2026-09-04T11:00:00Z");

        var ics = File.ReadAllText(outputPath);
        var start = ics.Split('\n').First(l => l.StartsWith("DTSTART", StringComparison.Ordinal)).Trim();
        Assert.Equal("DTSTART:20260904T100000Z", start);
    }

    [Fact]
    public void CreateAppointment_WithOffset_ShouldWriteTheSameInstantInUtc()
    {
        var tool = new EmailCalendarTool();
        var outputPath = CreateTestFilePath("offset.ics");

        tool.Execute("create", outputPath: outputPath,
            summary: "Review",
            startDate: "2026-09-04T12:00:00+02:00",
            endDate: "2026-09-04T13:00:00+02:00");

        var ics = File.ReadAllText(outputPath);
        var start = ics.Split('\n').First(l => l.StartsWith("DTSTART", StringComparison.Ordinal)).Trim();
        Assert.Equal("DTSTART:20260904T100000Z", start);
    }

    [Fact]
    public void CreateAppointment_WithoutZone_ShouldStillRoundTripLocally()
    {
        var tool = new EmailCalendarTool();
        var outputPath = CreateTestFilePath("naive.ics");

        tool.Execute("create", outputPath: outputPath,
            summary: "Review",
            startDate: "2026-09-04T10:00:00",
            endDate: "2026-09-04T11:00:00");

        var appointment = Appointment.Load(outputPath);
        Assert.Equal(new DateTime(2026, 9, 4, 10, 0, 0), appointment.StartDate);
    }
}
