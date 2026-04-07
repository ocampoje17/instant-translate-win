using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using InstantTranslateWin.App.Services;

namespace InstantTranslateWin.App;

public partial class App : System.Windows.Application
{
    private const string SingleInstanceMutexName = @"Local\InstantTranslateWin.App.SingleInstance";
    private const string ActivateEventName = @"Local\InstantTranslateWin.App.ActivateMainWindow";

    private Mutex? _singleInstanceMutex;
    private EventWaitHandle? _activateEvent;
    private CancellationTokenSource? _activateListenerCts;
    private Task? _activateListenerTask;

    public App()
    {
        // Avoid startup crashes on some machines where WPF D3D shader compilation fails.
        RenderOptions.ProcessRenderMode = System.Windows.Interop.RenderMode.SoftwareOnly;

        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        var createdNew = false;
        _singleInstanceMutex = new Mutex(true, SingleInstanceMutexName, out createdNew);
        if (!createdNew)
        {
            SignalRunningInstanceToActivate();
            Shutdown();
            return;
        }

        _activateEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ActivateEventName);
        StartActivateListener();

        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            _activateListenerCts?.Cancel();
            _activateEvent?.Set();
            _activateListenerTask?.Wait(TimeSpan.FromSeconds(1));
        }
        catch (Exception ex)
        {
            LogException("OnExit.SingleInstanceListenerShutdown", ex);
            // Ignore single-instance listener shutdown failures.
        }
        finally
        {
            _activateListenerCts?.Dispose();
            _activateListenerCts = null;

            _activateEvent?.Dispose();
            _activateEvent = null;

            if (_singleInstanceMutex is not null)
            {
                try
                {
                    _singleInstanceMutex.ReleaseMutex();
                }
                catch (ApplicationException)
                {
                    // Mutex may already be released during abnormal shutdown path.
                }

                _singleInstanceMutex.Dispose();
                _singleInstanceMutex = null;
            }
        }

        base.OnExit(e);
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        LogException("DispatcherUnhandledException", e.Exception);
        if (IsRecoverableUiException(e.Exception))
        {
            e.Handled = true;
            return;
        }

        e.Handled = true;
        _ = Dispatcher.BeginInvoke(() =>
        {
            try
            {
                MessageBox.Show(
                    "Ứng dụng vừa gặp lỗi nghiêm trọng và sẽ đóng để tránh trạng thái không ổn định.\nBạn có thể mở lại ngay sau đó.",
                    "Instant Translate",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
            catch (Exception ex)
            {
                LogException("OnDispatcherUnhandledException.ShowMessageBox", ex);
                // Ignore UI prompt failures.
            }

            try
            {
                Current?.Shutdown();
            }
            catch (Exception ex)
            {
                LogException("OnDispatcherUnhandledException.Shutdown", ex);
                // Ignore shutdown failures.
            }
        }, DispatcherPriority.Send);
    }

    private void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            LogException("AppDomainUnhandledException", ex);
            return;
        }

        LogRaw("AppDomainUnhandledException", e.ExceptionObject?.ToString() ?? "Unknown exception object");
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        LogException("UnobservedTaskException", e.Exception);
        e.SetObserved();
    }

    private static void SignalRunningInstanceToActivate()
    {
        const int maxAttempts = 10;
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            try
            {
                using var activateEvent = EventWaitHandle.OpenExisting(ActivateEventName);
                activateEvent.Set();
                return;
            }
            catch (WaitHandleCannotBeOpenedException) when (attempt < maxAttempts - 1)
            {
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                LogException("SignalRunningInstanceToActivate", ex);
                return;
            }
        }
    }

    private void StartActivateListener()
    {
        if (_activateEvent is null)
        {
            return;
        }

        _activateListenerCts = new CancellationTokenSource();
        var cancellationToken = _activateListenerCts.Token;
        var activateEvent = _activateEvent;
        _activateListenerTask = Task.Run(() => ListenForActivateSignal(activateEvent, cancellationToken), cancellationToken);
    }

    private void ListenForActivateSignal(EventWaitHandle activateEvent, CancellationToken cancellationToken)
    {
        var waitHandles = new WaitHandle[] { activateEvent, cancellationToken.WaitHandle };
        while (!cancellationToken.IsCancellationRequested)
        {
            var waitResult = WaitHandle.WaitAny(waitHandles);
            if (waitResult == 1)
            {
                break;
            }

            _ = Dispatcher.BeginInvoke(RestoreMainWindowFromExternalLaunch, DispatcherPriority.Send);
        }
    }

    private void RestoreMainWindowFromExternalLaunch()
    {
        var mainWindow = MainWindow;
        if (mainWindow is null)
        {
            return;
        }

        try
        {
            if (!mainWindow.IsVisible)
            {
                mainWindow.Show();
            }

            mainWindow.ShowInTaskbar = true;
            mainWindow.WindowState = WindowState.Normal;
            mainWindow.Activate();
            mainWindow.Topmost = true;
            mainWindow.Topmost = false;
            mainWindow.Focus();
        }
        catch (Exception ex)
        {
            LogException("RestoreMainWindowFromExternalLaunch", ex);
        }
    }

    private static void LogException(string source, Exception exception)
    {
        ErrorFileLogger.LogException(source, exception);
    }

    private static bool IsRecoverableUiException(Exception ex)
    {
        if (ex is OperationCanceledException or TaskCanceledException)
        {
            return true;
        }

        if (ex is COMException comEx)
        {
            const int clipbrdCantOpen = unchecked((int)0x800401D0);
            const int clipbrdCantEmpty = unchecked((int)0x800401D1);
            const int rpcCallRejected = unchecked((int)0x80010001);
            return comEx.HResult == clipbrdCantOpen ||
                   comEx.HResult == clipbrdCantEmpty ||
                   comEx.HResult == rpcCallRejected;
        }

        return false;
    }

    private static void LogRaw(string source, string message)
    {
        ErrorFileLogger.LogMessage(source, message);
    }
}
