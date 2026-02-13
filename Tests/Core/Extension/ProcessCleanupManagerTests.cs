using System.Diagnostics;
using AsposeMcpServer.Core.Extension;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for ProcessCleanupManager class.
/// </summary>
public class ProcessCleanupManagerTests : IDisposable
{
    private readonly Mock<ILogger> _loggerMock = new();
    private readonly List<Process> _testProcesses = new();
    private ProcessCleanupManager? _manager;

    public void Dispose()
    {
        _manager?.Dispose();

        foreach (var process in _testProcesses)
            try
            {
                if (!process.HasExited)
                    process.Kill();
                process.Dispose();
            }
            // ReSharper disable once EmptyGeneralCatchClause - Best-effort process cleanup in dispose
            catch
            {
            }

        _testProcesses.Clear();
        GC.SuppressFinalize(this);
    }

    #region Process Auto-Unregister Tests

    [Fact]
    public void RegisteredProcess_WhenExits_AutomaticallyUnregistered()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        _manager.RegisterProcess(process);

        process.Kill();
        process.WaitForExit();

        Thread.Sleep(100);

        var zombies = _manager.GetZombieProcesses();
        Assert.DoesNotContain(process.Id, zombies);
    }

    #endregion

    #region Constructor Tests

    [Fact]
    public void Constructor_WithoutLogger_DoesNotThrow()
    {
        var exception = Record.Exception(() => { _manager = new ProcessCleanupManager(); });

        Assert.Null(exception);
    }

    [Fact]
    public void Constructor_WithLogger_DoesNotThrow()
    {
        var exception = Record.Exception(() => { _manager = new ProcessCleanupManager(_loggerMock.Object); });

        Assert.Null(exception);
    }

    #endregion

    #region RegisterProcess Tests

    [Fact]
    public void RegisterProcess_AfterDispose_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        _manager.Dispose();
        var process = CreateTestProcess();

        var exception = Record.Exception(() => _manager.RegisterProcess(process));

        Assert.Null(exception);
    }

    [Fact]
    public void RegisterProcess_ExitedProcess_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        process.Kill();
        process.WaitForExit();

        var exception = Record.Exception(() => _manager.RegisterProcess(process));

        Assert.Null(exception);
    }

    [Fact]
    public void RegisterProcess_RunningProcess_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();

        var exception = Record.Exception(() => _manager.RegisterProcess(process));

        Assert.Null(exception);
    }

    [Fact]
    public void RegisterProcess_MultipleProcesses_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process1 = CreateTestProcess();
        var process2 = CreateTestProcess();

        var exception = Record.Exception(() =>
        {
            _manager.RegisterProcess(process1);
            _manager.RegisterProcess(process2);
        });

        Assert.Null(exception);
    }

    [Fact]
    public void RegisterProcess_TrackedProcess_IsKilledOnDispose()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateLongRunningProcess();

        _manager.RegisterProcess(process);
        _manager.Dispose();
        _manager = null;

        process.WaitForExit(1000);
        Assert.True(process.HasExited);
    }

    #endregion

    #region UnregisterProcess Tests

    [Fact]
    public void UnregisterProcess_RegisteredProcess_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        _manager.RegisterProcess(process);

        var exception = Record.Exception(() => _manager.UnregisterProcess(process));

        Assert.Null(exception);
    }

    [Fact]
    public void UnregisterProcess_UnregisteredProcess_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();

        var exception = Record.Exception(() => _manager.UnregisterProcess(process));

        Assert.Null(exception);
    }

    [Fact]
    public void UnregisterProcess_AfterDispose_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        _manager.RegisterProcess(process);
        _manager.Dispose();

        var exception = Record.Exception(() => _manager.UnregisterProcess(process));

        Assert.Null(exception);
    }

    #endregion

    #region IsZombieProcess Tests

    [Fact]
    public void IsZombieProcess_ExitedProcess_ReturnsFalse()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        process.Kill();
        process.WaitForExit();

        var result = _manager.IsZombieProcess(process);

        Assert.False(result);
    }

    [Fact]
    public void IsZombieProcess_RunningProcess_ReturnsFalse()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateLongRunningProcess();

        var result = _manager.IsZombieProcess(process);

        Assert.False(result);
    }

    [Fact]
    public void IsZombieProcess_QuicklyExitingProcess_ReturnsFalse()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        process.WaitForExit(1000);

        var result = _manager.IsZombieProcess(process);

        Assert.False(result);
    }

    #endregion

    #region GetZombieProcesses Tests

    [Fact]
    public void GetZombieProcesses_NoTrackedProcesses_ReturnsEmptyList()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);

        var result = _manager.GetZombieProcesses();

        Assert.Empty(result);
    }

    [Fact]
    public void GetZombieProcesses_AllRunningProcesses_ReturnsEmptyList()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process1 = CreateTestProcess();
        var process2 = CreateTestProcess();
        _manager.RegisterProcess(process1);
        _manager.RegisterProcess(process2);

        var result = _manager.GetZombieProcesses();

        Assert.Empty(result);
    }

    [Fact]
    public void GetZombieProcesses_ReturnsReadOnlyList()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);

        var result = _manager.GetZombieProcesses();

        Assert.IsType<IReadOnlyList<int>>(result, false);
    }

    #endregion

    #region Dispose Tests

    [Fact]
    public void Dispose_MultipleTimes_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);

        _manager.Dispose();
        var exception = Record.Exception(() => _manager.Dispose());

        Assert.Null(exception);
    }

    [Fact]
    public void Dispose_KillsTrackedProcesses()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateLongRunningProcess();
        _manager.RegisterProcess(process);

        _manager.Dispose();

        process.WaitForExit(1000);
        Assert.True(process.HasExited);
    }

    [Fact]
    public void Dispose_WithExitedProcesses_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        _manager.RegisterProcess(process);
        process.Kill();
        process.WaitForExit();

        var exception = Record.Exception(() => _manager.Dispose());

        Assert.Null(exception);
    }

    [Fact]
    public void Dispose_WithMixedProcesses_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var runningProcess = CreateLongRunningProcess();
        var exitedProcess = CreateTestProcess();
        _manager.RegisterProcess(runningProcess);
        _manager.RegisterProcess(exitedProcess);
        exitedProcess.Kill();
        exitedProcess.WaitForExit();

        var exception = Record.Exception(() => _manager.Dispose());

        Assert.Null(exception);
    }

    #endregion

    #region Edge Case Tests

    [Fact]
    public void RegisterProcess_SameProcessTwice_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();

        _manager.RegisterProcess(process);
        var exception = Record.Exception(() => _manager.RegisterProcess(process));

        Assert.Null(exception);
    }

    [Fact]
    public void UnregisterProcess_SameProcessTwice_DoesNotThrow()
    {
        _manager = new ProcessCleanupManager(_loggerMock.Object);
        var process = CreateTestProcess();
        _manager.RegisterProcess(process);

        _manager.UnregisterProcess(process);
        var exception = Record.Exception(() => _manager.UnregisterProcess(process));

        Assert.Null(exception);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test process that exits quickly.
    /// </summary>
    private Process CreateTestProcess()
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = OperatingSystem.IsWindows() ? "cmd.exe" : "/bin/sh",
            Arguments = OperatingSystem.IsWindows() ? "/c echo test" : "-c \"echo test\"",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true
        };

        var process = Process.Start(startInfo)!;
        _testProcesses.Add(process);
        return process;
    }

    /// <summary>
    ///     Creates a test process that runs for a longer time.
    /// </summary>
    private Process CreateLongRunningProcess()
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = OperatingSystem.IsWindows() ? "cmd.exe" : "/bin/sh",
            Arguments = OperatingSystem.IsWindows() ? "/c ping -n 60 127.0.0.1" : "-c \"sleep 60\"",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true
        };

        var process = Process.Start(startInfo)!;
        _testProcesses.Add(process);
        return process;
    }

    #endregion
}
