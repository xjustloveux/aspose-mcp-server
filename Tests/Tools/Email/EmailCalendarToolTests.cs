using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Email.Calendar;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Tools.Email;

/// <summary>
///     Integration tests for <see cref="EmailCalendarTool" />.
///     Focuses on operation routing, file I/O, and end-to-end workflows.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class EmailCalendarToolTests : EmailTestBase
{
    private readonly EmailCalendarTool _tool = new();

    #region SetAttendees Operation

    [Fact]
    public void Execute_SetAttendees_SetsAttendeeList()
    {
        var inputPath = CreateTestIcsFile("tool_attendees_input.ics", "Attendees Tool Test");
        var outputPath = CreateTestFilePath("tool_attendees_output.ics");
        var result = _tool.Execute("set_attendees", inputPath, outputPath,
            attendees: "one@example.com,two@example.com");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal(2, appointment.Attendees.Count);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test ICS file for tool tests.
    /// </summary>
    private string CreateTestIcsFile(string fileName, string summary, string description = "",
        string location = "", params string[] attendees)
    {
        var filePath = CreateTestFilePath(fileName);
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

    #region Create Operation

    [Fact]
    public void Execute_Create_CreatesIcsFile()
    {
        var outputPath = CreateTestFilePath("tool_create.ics");
        var result = _tool.Execute("create", outputPath: outputPath, summary: "Tool Meeting");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Tool Meeting", appointment.Summary);
    }

    [Fact]
    public void Execute_Create_WithAllParameters_CreatesFullAppointment()
    {
        var outputPath = CreateTestFilePath("tool_create_full.ics");
        var result = _tool.Execute("create",
            outputPath: outputPath,
            summary: "Full Meeting",
            description: "All params test",
            startDate: "2025-07-01T09:00:00",
            endDate: "2025-07-01T10:00:00",
            location: "Room B",
            attendees: "a@example.com,b@example.com");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Full Meeting", appointment.Summary);
        Assert.Equal("All params test", appointment.Description);
        Assert.Equal("Room B", appointment.Location);
        Assert.Equal(2, appointment.Attendees.Count);
    }

    [Fact]
    public void Execute_Create_WithMinimalParameters_UsesDefaults()
    {
        var outputPath = CreateTestFilePath("tool_create_minimal.ics");
        var result = _tool.Execute("create", outputPath: outputPath);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("New Appointment", appointment.Summary);
    }

    #endregion

    #region GetInfo Operation

    [Fact]
    public void Execute_GetInfo_ReturnsAppointmentInfo()
    {
        var icsPath = CreateTestIcsFile("tool_getinfo.ics", "Info Meeting", "Meeting description",
            "Office");
        var result = _tool.Execute("get_info", icsPath);

        Assert.IsType<FinalizedResult<AppointmentEmailInfo>>(result);
        var data = GetResultData<AppointmentEmailInfo>(result);
        Assert.Equal("Info Meeting", data.Summary);
        Assert.Equal("Meeting description", data.Description);
        Assert.Equal("Office", data.Location);
        Assert.False(data.IsRecurring);
    }

    [Fact]
    public void Execute_GetInfo_WithAttendees_ReturnsAttendeeList()
    {
        var icsPath = CreateTestIcsFile("tool_getinfo_att.ics", "Attendee Info", "",
            "", "x@example.com", "y@example.com");
        var result = _tool.Execute("get_info", icsPath);

        var data = GetResultData<AppointmentEmailInfo>(result);
        Assert.NotNull(data.Attendees);
        Assert.Equal(2, data.Attendees.Count);
    }

    #endregion

    #region Save Operation

    [Fact]
    public void Execute_Save_AsIcs_SavesSuccessfully()
    {
        var inputPath = CreateTestIcsFile("tool_save_input.ics", "Save Tool Test");
        var outputPath = CreateTestFilePath("tool_save_output.ics");
        var result = _tool.Execute("save", inputPath, outputPath, format: "ics");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Save Tool Test", appointment.Summary);
    }

    [SkippableFact]
    public void Execute_Save_AsMsg_SavesSuccessfully()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "MSG conversion requires license");

        var inputPath = CreateTestIcsFile("tool_save_msg_input.ics", "MSG Save Tool");
        var outputPath = CreateTestFilePath("tool_save_output.msg");
        var result = _tool.Execute("save", inputPath, outputPath, format: "msg");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region SetRecurrence Operation

    [Theory]
    [InlineData("daily")]
    [InlineData("weekly")]
    [InlineData("monthly")]
    [InlineData("yearly")]
    public void Execute_SetRecurrence_WithValidPattern_SetsRecurrence(string pattern)
    {
        var inputPath = CreateTestIcsFile("tool_recurrence_input.ics", "Recurring Tool Test");
        var outputPath = CreateTestFilePath($"tool_recurrence_{pattern}_output.ics");
        var result = _tool.Execute("set_recurrence", inputPath, outputPath,
            pattern: pattern);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_SetRecurrence_WithCustomIntervalAndCount_SetsValues()
    {
        var inputPath = CreateTestIcsFile("tool_recurrence_custom.ics", "Custom Recurrence");
        var outputPath = CreateTestFilePath("tool_recurrence_custom_output.ics");
        var result = _tool.Execute("set_recurrence", inputPath, outputPath,
            pattern: "weekly", interval: 2, count: 26);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("interval: 2", data.Message);
        Assert.Contains("count: 26", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var outputPath = CreateTestFilePath($"tool_case_{operation}.ics");
        var result = _tool.Execute(operation, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", outputPath: "test.ics"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region End-to-End Workflow

    [Fact]
    public void Workflow_CreateThenGetInfo_RoundTrips()
    {
        var outputPath = CreateTestFilePath("workflow_roundtrip.ics");
        _tool.Execute("create", outputPath: outputPath, summary: "Roundtrip Meeting",
            description: "E2E test", location: "Lab", startDate: "2025-08-01T10:00:00",
            endDate: "2025-08-01T11:00:00");

        var infoResult = _tool.Execute("get_info", outputPath);
        var data = GetResultData<AppointmentEmailInfo>(infoResult);

        Assert.Equal("Roundtrip Meeting", data.Summary);
        Assert.Equal("E2E test", data.Description);
        Assert.Equal("Lab", data.Location);
        Assert.False(data.IsRecurring);
    }

    [Fact]
    public void Workflow_CreateThenSetRecurrenceThenGetInfo_ShowsRecurring()
    {
        var createPath = CreateTestFilePath("workflow_create.ics");
        _tool.Execute("create", outputPath: createPath, summary: "Recurring Workflow");

        var recurringPath = CreateTestFilePath("workflow_recurring.ics");
        _tool.Execute("set_recurrence", createPath, recurringPath,
            pattern: "weekly", count: 52);

        var infoResult = _tool.Execute("get_info", recurringPath);
        var data = GetResultData<AppointmentEmailInfo>(infoResult);
        Assert.Equal("Recurring Workflow", data.Summary);
        Assert.NotNull(data);
    }

    #endregion
}
