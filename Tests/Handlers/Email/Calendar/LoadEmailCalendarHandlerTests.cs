using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Handlers.Email.Calendar;
using AsposeMcpServer.Results.Email.Calendar;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Calendar;

/// <summary>
///     Tests for <see cref="LoadEmailCalendarHandler" />.
/// </summary>
public class LoadEmailCalendarHandlerTests : HandlerTestBase<object>
{
    private readonly LoadEmailCalendarHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetInfo()
    {
        Assert.Equal("get_info", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test ICS file with the specified properties.
    /// </summary>
    private string CreateTestIcsFile(string fileName, string summary, string description = "",
        string location = "", params string[] attendees)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var attendeeCollection = new MailAddressCollection();
        foreach (var attendee in attendees)
            attendeeCollection.Add(new MailAddress(attendee));

        var appointment = new Appointment(
            location,
            summary,
            description,
            DateTime.Now,
            DateTime.Now.AddHours(1),
            new MailAddress("organizer@example.com"),
            attendeeCollection);
        appointment.Save(filePath, AppointmentSaveFormat.Ics);
        return filePath;
    }

    #endregion

    #region Basic Load Operations

    [Fact]
    public void Execute_WithValidIcsFile_ReturnsAppointmentInfo()
    {
        var icsPath = CreateTestIcsFile("load_test.ics", "Project Review", "Review Q3 progress",
            "Main Office", "reviewer@example.com");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", icsPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<AppointmentEmailInfo>(res);
        Assert.Equal("Project Review", result.Summary);
        Assert.Equal("Review Q3 progress", result.Description);
        Assert.Equal("Main Office", result.Location);
        Assert.NotNull(result.StartDate);
        Assert.NotNull(result.EndDate);
        Assert.Contains("Project Review", result.Message);
    }

    [Fact]
    public void Execute_WithAttendees_ReturnsAttendeeList()
    {
        var icsPath = CreateTestIcsFile("attendees_test.ics", "Team Meeting", "Sync up",
            "Room A", "alice@example.com", "bob@example.com");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", icsPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<AppointmentEmailInfo>(res);
        Assert.NotNull(result.Attendees);
        Assert.Equal(2, result.Attendees.Count);
        Assert.Contains("alice@example.com", result.Attendees);
        Assert.Contains("bob@example.com", result.Attendees);
    }

    [Fact]
    public void Execute_WithNoAttendees_ReturnsNullAttendees()
    {
        var icsPath = Path.Combine(TestDir, "no_attendees.ics");
        var appointment = new Appointment(
            "Office",
            "Solo Meeting",
            "Description",
            DateTime.Now,
            DateTime.Now.AddHours(1),
            new MailAddress("organizer@example.com"),
            new MailAddressCollection());
        appointment.Save(icsPath, AppointmentSaveFormat.Ics);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", icsPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<AppointmentEmailInfo>(res);
        Assert.Null(result.Attendees);
    }

    [Fact]
    public void Execute_WithNonRecurringAppointment_ReturnsIsRecurringFalse()
    {
        var icsPath = CreateTestIcsFile("non_recurring.ics", "One-off Meeting");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", icsPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<AppointmentEmailInfo>(res);
        Assert.False(result.IsRecurring);
    }

    [Fact]
    public void Execute_ReturnsIso8601DateFormats()
    {
        var icsPath = CreateTestIcsFile("date_format.ics", "Date Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", icsPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<AppointmentEmailInfo>(res);
        Assert.NotNull(result.StartDate);
        Assert.NotNull(result.EndDate);
        Assert.True(DateTime.TryParse(result.StartDate, out _), "StartDate should be a valid date");
        Assert.True(DateTime.TryParse(result.EndDate, out _), "EndDate should be a valid date");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.ics") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
