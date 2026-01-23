using System.Text.Json;
using AsposeMcpServer.Core.Tasks;
using TaskStatus = AsposeMcpServer.Core.Tasks.TaskStatus;

namespace AsposeMcpServer.Tests.Core.Tasks;

public class TaskStoreTests : IDisposable
{
    private readonly TaskConfig _config;
    private readonly TaskStore _store;

    public TaskStoreTests()
    {
        _config = new TaskConfig
        {
            MaxConcurrentTasks = 3,
            DefaultTtlMs = 60000,
            MaxTtlMs = 300000
        };
        _store = new TaskStore(_config);
    }

    public void Dispose()
    {
        _store.Clear();
    }

    [Fact]
    public void Constructor_WithNullConfig_ShouldThrow()
    {
        Assert.Throws<ArgumentNullException>(() => new TaskStore(null!));
    }

    [Fact]
    public void CreateTask_ShouldReturnTaskWithUniqueId()
    {
        var args = JsonSerializer.SerializeToElement(new { path = "test.docx" });

        var task1 = _store.CreateTask("convert_to_pdf", args);
        var task2 = _store.CreateTask("convert_to_pdf", args);

        Assert.NotEqual(task1.TaskId, task2.TaskId);
        Assert.Equal(32, task1.TaskId.Length); // GUID without dashes
    }

    [Fact]
    public void CreateTask_ShouldSetCorrectInitialValues()
    {
        var args = JsonSerializer.SerializeToElement(new { path = "test.docx" });

        var task = _store.CreateTask("convert_to_pdf", args, 120000, "user1");

        Assert.Equal("convert_to_pdf", task.ToolName);
        Assert.Equal(TaskStatus.Working, task.Status);
        Assert.Equal(120000, task.Ttl);
        Assert.Equal("user1", task.OwnerId);
        Assert.NotNull(task.StatusMessage);
    }

    [Fact]
    public void CreateTask_WithNullToolName_ShouldThrow()
    {
        var args = JsonSerializer.SerializeToElement(new { });

        Assert.Throws<ArgumentNullException>(() => _store.CreateTask(null!, args));
    }

    [Fact]
    public void CreateTask_ShouldEnforceTtlLimit()
    {
        var args = JsonSerializer.SerializeToElement(new { });

        var task = _store.CreateTask("convert_to_pdf", args, 999999999);

        Assert.Equal(_config.MaxTtlMs, task.Ttl);
    }

    [Fact]
    public void CreateTask_WhenMaxConcurrentReached_ShouldThrow()
    {
        var args = JsonSerializer.SerializeToElement(new { });

        _store.CreateTask("tool1", args);
        _store.CreateTask("tool2", args);
        _store.CreateTask("tool3", args);

        Assert.Throws<InvalidOperationException>(() => _store.CreateTask("tool4", args));
    }

    [Fact]
    public void CreateTask_WithOwnerIsolation_ShouldEnforceLimitPerOwner()
    {
        var args = JsonSerializer.SerializeToElement(new { });

        _store.CreateTask("tool1", args, ownerId: "user1");
        _store.CreateTask("tool2", args, ownerId: "user1");
        _store.CreateTask("tool3", args, ownerId: "user1");

        // Different owner should be able to create
        var task = _store.CreateTask("tool4", args, ownerId: "user2");
        Assert.NotNull(task);

        // Same owner should fail
        Assert.Throws<InvalidOperationException>(() => _store.CreateTask("tool5", args, ownerId: "user1"));
    }

    [Fact]
    public void GetTask_WithValidId_ShouldReturnTask()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var created = _store.CreateTask("convert_to_pdf", args);

        var retrieved = _store.GetTask(created.TaskId);

        Assert.NotNull(retrieved);
        Assert.Equal(created.TaskId, retrieved.TaskId);
    }

    [Fact]
    public void GetTask_WithInvalidId_ShouldReturnNull()
    {
        var result = _store.GetTask("nonexistent");

        Assert.Null(result);
    }

    [Fact]
    public void GetTask_WithNullId_ShouldReturnNull()
    {
        var result = _store.GetTask(null!);

        Assert.Null(result);
    }

    [Fact]
    public void GetTask_WithOwnerMismatch_ShouldReturnNull()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task = _store.CreateTask("tool", args, ownerId: "user1");

        var result = _store.GetTask(task.TaskId, "user2");

        Assert.Null(result);
    }

    [Fact]
    public void GetTask_WithMatchingOwner_ShouldReturnTask()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task = _store.CreateTask("tool", args, ownerId: "user1");

        var result = _store.GetTask(task.TaskId, "user1");

        Assert.NotNull(result);
    }

    [Fact]
    public void ListTasks_ShouldReturnAllTasks()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        _store.CreateTask("tool1", args);
        _store.CreateTask("tool2", args);

        var tasks = _store.ListTasks();

        Assert.Equal(2, tasks.Count);
    }

    [Fact]
    public void ListTasks_WithOwnerId_ShouldFilterByOwner()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        _store.CreateTask("tool1", args, ownerId: "user1");
        _store.CreateTask("tool2", args, ownerId: "user1");
        _store.CreateTask("tool3", args, ownerId: "user2");

        var tasks = _store.ListTasks("user1");

        Assert.Equal(2, tasks.Count);
        Assert.All(tasks, t => Assert.Equal("user1", t.OwnerId));
    }

    [Fact]
    public void ListTasks_ShouldReturnNewestFirst()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task1 = _store.CreateTask("tool1", args);
        Thread.Sleep(10);
        var task2 = _store.CreateTask("tool2", args);

        var tasks = _store.ListTasks();

        Assert.Equal(task2.TaskId, tasks[0].TaskId);
        Assert.Equal(task1.TaskId, tasks[1].TaskId);
    }

    [Fact]
    public void UpdateTaskStatus_ShouldUpdateStatus()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task = _store.CreateTask("tool", args);

        var result = _store.UpdateTaskStatus(task.TaskId, TaskStatus.Completed, "Done", "Success");

        Assert.True(result);
        Assert.Equal(TaskStatus.Completed, task.Status);
        Assert.Equal("Done", task.StatusMessage);
        Assert.Equal("Success", task.Result);
    }

    [Fact]
    public void UpdateTaskStatus_WithInvalidId_ShouldReturnFalse()
    {
        var result = _store.UpdateTaskStatus("nonexistent", TaskStatus.Completed);

        Assert.False(result);
    }

    [Fact]
    public void UpdateTaskStatus_WithError_ShouldSetErrorMessage()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task = _store.CreateTask("tool", args);

        _store.UpdateTaskStatus(task.TaskId, TaskStatus.Failed, "Failed", errorMessage: "File not found");

        Assert.Equal(TaskStatus.Failed, task.Status);
        Assert.Equal("File not found", task.ErrorMessage);
    }

    [Fact]
    public void CancelTask_ShouldCancelWorkingTask()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task = _store.CreateTask("tool", args);

        var result = _store.CancelTask(task.TaskId);

        Assert.True(result);
        Assert.Equal(TaskStatus.Cancelled, task.Status);
    }

    [Fact]
    public void CancelTask_WithInvalidId_ShouldReturnFalse()
    {
        var result = _store.CancelTask("nonexistent");

        Assert.False(result);
    }

    [Fact]
    public void CancelTask_WithTerminalTask_ShouldReturnFalse()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task = _store.CreateTask("tool", args);
        _store.UpdateTaskStatus(task.TaskId, TaskStatus.Completed);

        var result = _store.CancelTask(task.TaskId);

        Assert.False(result);
    }

    [Fact]
    public void CancelTask_WithOwnerMismatch_ShouldReturnFalse()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        var task = _store.CreateTask("tool", args, ownerId: "user1");

        var result = _store.CancelTask(task.TaskId, "user2");

        Assert.False(result);
    }

    [Fact]
    public void CleanupExpiredTasks_ShouldRemoveExpiredTerminalTasks()
    {
        var shortTtlConfig = new TaskConfig
        {
            MaxConcurrentTasks = 10,
            DefaultTtlMs = 1
        };
        var store = new TaskStore(shortTtlConfig);
        var args = JsonSerializer.SerializeToElement(new { });

        var task = store.CreateTask("tool", args);
        store.UpdateTaskStatus(task.TaskId, TaskStatus.Completed);

        Thread.Sleep(10);

        var removed = store.CleanupExpiredTasks();

        Assert.Equal(1, removed);
        Assert.Null(store.GetTask(task.TaskId));
    }

    [Fact]
    public void CleanupExpiredTasks_ShouldNotRemoveWorkingTasks()
    {
        var shortTtlConfig = new TaskConfig
        {
            MaxConcurrentTasks = 10,
            DefaultTtlMs = 1
        };
        var store = new TaskStore(shortTtlConfig);
        var args = JsonSerializer.SerializeToElement(new { });

        var task = store.CreateTask("tool", args);
        Thread.Sleep(10);

        var removed = store.CleanupExpiredTasks();

        Assert.Equal(0, removed);
        Assert.NotNull(store.GetTask(task.TaskId));
    }

    [Fact]
    public void ActiveTaskCount_ShouldCountOnlyWorkingTasks()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        _store.CreateTask("tool1", args);
        var task2 = _store.CreateTask("tool2", args);
        _store.UpdateTaskStatus(task2.TaskId, TaskStatus.Completed);

        Assert.Equal(1, _store.ActiveTaskCount);
    }

    [Fact]
    public void TotalTaskCount_ShouldCountAllTasks()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        _store.CreateTask("tool1", args);
        var task2 = _store.CreateTask("tool2", args);
        _store.UpdateTaskStatus(task2.TaskId, TaskStatus.Completed);

        Assert.Equal(2, _store.TotalTaskCount);
    }

    [Fact]
    public void Clear_ShouldRemoveAllTasks()
    {
        var args = JsonSerializer.SerializeToElement(new { });
        _store.CreateTask("tool1", args);
        _store.CreateTask("tool2", args);

        _store.Clear();

        Assert.Equal(0, _store.TotalTaskCount);
    }
}
