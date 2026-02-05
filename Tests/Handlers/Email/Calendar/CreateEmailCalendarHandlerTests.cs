using Aspose.Email.Calendar;
using AsposeMcpServer.Handlers.Email.Calendar;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Calendar;

/// <summary>
///     Tests for <see cref="CreateEmailCalendarHandler" />.
/// </summary>
public class CreateEmailCalendarHandlerTests : HandlerTestBase<object>
{
    private readonly CreateEmailCalendarHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_WithOutputPathOnly_CreatesDefaultAppointment()
    {
        var outputPath = Path.Combine(TestDir, "default.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("New Appointment", result.Message);
        Assert.Contains(outputPath, result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("New Appointment", appointment.Summary);
    }

    [Fact]
    public void Execute_WithAllParameters_CreatesFullAppointment()
    {
        var outputPath = Path.Combine(TestDir, "full.ics");
        var context = CreateContext(new object());
        var startDate = "2025-06-15T10:00:00";
        var endDate = "2025-06-15T11:30:00";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "summary", "Team Meeting" },
            { "description", "Weekly sync" },
            { "startDate", startDate },
            { "endDate", endDate },
            { "location", "Conference Room A" },
            { "attendees", "alice@example.com,bob@example.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Team Meeting", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Team Meeting", appointment.Summary);
        Assert.Equal("Weekly sync", appointment.Description);
        Assert.Equal("Conference Room A", appointment.Location);
        Assert.Equal(2, appointment.Attendees.Count);
    }

    [Fact]
    public void Execute_WithSummaryOnly_CreatesAppointmentWithSummary()
    {
        var outputPath = Path.Combine(TestDir, "summary_only.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "summary", "Quick Standup" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Quick Standup", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Quick Standup", appointment.Summary);
    }

    [Fact]
    public void Execute_WithStartDateOnly_SetsEndDateOneHourLater()
    {
        var outputPath = Path.Combine(TestDir, "start_only.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "startDate", "2025-06-15T14:00:00" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        var duration = appointment.EndDate - appointment.StartDate;
        Assert.Equal(TimeSpan.FromHours(1), duration);
    }

    [Fact]
    public void Execute_WithLocation_SetsLocation()
    {
        var outputPath = Path.Combine(TestDir, "with_location.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "location", "Building 42, Room 101" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Building 42, Room 101", appointment.Location);
    }

    [Fact]
    public void Execute_WithDescription_SetsDescription()
    {
        var outputPath = Path.Combine(TestDir, "with_desc.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "description", "Discuss Q3 planning" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Discuss Q3 planning", appointment.Description);
    }

    [Fact]
    public void Execute_WithSingleAttendee_AddsOneAttendee()
    {
        var outputPath = Path.Combine(TestDir, "single_attendee.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "attendees", "solo@example.com" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Single(appointment.Attendees);
    }

    [Fact]
    public void Execute_WithMultipleAttendees_AddsAllAttendees()
    {
        var outputPath = Path.Combine(TestDir, "multi_attendees.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "attendees", "a@example.com, b@example.com, c@example.com" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var appointment = Appointment.Load(outputPath);
        Assert.Equal(3, appointment.Attendees.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidStartDate_ThrowsFormatException()
    {
        var outputPath = Path.Combine(TestDir, "bad_date.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "startDate", "not-a-date" }
        });

        Assert.Throws<FormatException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidEndDate_ThrowsFormatException()
    {
        var outputPath = Path.Combine(TestDir, "bad_end_date.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "startDate", "2025-06-15T10:00:00" },
            { "endDate", "invalid-end" }
        });

        Assert.Throws<FormatException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
