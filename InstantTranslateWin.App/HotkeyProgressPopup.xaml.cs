using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using InstantTranslateWin.App.Services;
using Wpf.Ui.Controls;

namespace InstantTranslateWin.App;

public partial class HotkeyProgressPopup : FluentWindow
{
    public event EventHandler? RestartAppRequested;
    public event EventHandler? QuickInputRequested;

    private const uint MonitorDefaultToNearest = 0x00000002;
    private const double EdgeMargin = 16;

    private CancellationTokenSource? _closeCts;
    private bool _isPointerHovering;
    private bool _autoCloseRequested;
    private bool _isUsingCustomPosition;
    private TimeSpan _scheduledCloseDelay = TimeSpan.FromSeconds(3);

    private readonly SolidColorBrush _successBadgeBrush = new(Color.FromRgb(54, 158, 94));
    private readonly SolidColorBrush _errorBadgeBrush = new(Color.FromRgb(216, 76, 76));
    private readonly SolidColorBrush _infoBadgeBrush = new(Color.FromRgb(70, 125, 204));
    private readonly SolidColorBrush _processingBadgeBrush = new(Color.FromRgb(210, 138, 38));
    private readonly SolidColorBrush _clipboardBadgeBrush = new(Color.FromRgb(45, 152, 146));
    private readonly SolidColorBrush _historyBadgeBrush = new(Color.FromRgb(113, 91, 186));
    private readonly HotkeyProgressPopupPositionStore _positionStore = new();

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMonitorInfo
    {
        public int Size;
        public NativeRect Monitor;
        public NativeRect Work;
        public uint Flags;
    }

    private delegate bool NativeMonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, IntPtr lprcMonitor, IntPtr dwData);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromPoint(NativePoint pt, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref NativeMonitorInfo lpmi);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumDisplayMonitors(
        IntPtr hdc,
        IntPtr lprcClip,
        NativeMonitorEnumProc lpfnEnum,
        IntPtr dwData
    );

    public HotkeyProgressPopup()
    {
        InitializeComponent();
    }

    public void ShowAtBottomRight()
    {
        if (!IsVisible)
        {
            Show();
            UpdateLayout();
        }

        try
        {
            PositionAtPreferredLocation();
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopup.ShowAtBottomRight.PositionFallback", ex);
            ApplyBottomRightPosition();
        }
    }

    public void BeginProgress(string title, string subtitle)
    {
        CancelScheduledClose();
        _autoCloseRequested = false;

        TitleTextBlock.Text = title;
        SetCurrentStep("Khởi tạo...", SymbolRegular.Rocket24);
        SetSourcePreview(subtitle);

        LoadingPage.Visibility = Visibility.Visible;
        ResultPage.Visibility = Visibility.Hidden;
        ResultActionsPanel.Visibility = Visibility.Collapsed;
        WorkingProgressRing.IsIndeterminate = true;
        WorkingProgressRing.Visibility = Visibility.Visible;
    }

    public void AddStep(string step)
    {
        var normalizedStep = NormalizeSingleLine(step, 80);
        var icon = ResolveStepIcon(normalizedStep);
        SetCurrentStep(normalizedStep, icon);
    }

    public void SetProgressState(string subtitle, double? progressPercent = null)
    {
        _ = progressPercent;
        var normalizedStep = NormalizeSingleLine(subtitle, 80);
        SetCurrentStep(normalizedStep, ResolveStepIcon(normalizedStep));
    }

    public void SetSourcePreview(string sourceText)
    {
        SourcePreviewTextBlock.Text = $"Nguồn: {NormalizeSingleLine(sourceText, 92)}";
    }

    public void SetSuccessState(string subtitle, string resultText)
    {
        LoadingPage.Visibility = Visibility.Hidden;
        ResultPage.Visibility = Visibility.Visible;
        ResultActionsPanel.Visibility = Visibility.Collapsed;

        ResultStateIcon.Symbol = SymbolRegular.Checkmark24;
        ResultIconBadge.Background = _successBadgeBrush;
        ResultStateIcon.Foreground = new SolidColorBrush(Colors.White);

        ResultTitleTextBlock.Text = "Dịch thành công";
        ResultMessageTextBlock.Text = $"Kết quả: {NormalizeSingleLine(resultText, 96)}";
        TitleTextBlock.Text = NormalizeSingleLine(subtitle, 64);
        EnsureCurrentPositionIsVisible();
    }

    public void SetErrorState(string subtitle, bool showQuickInputAction = false, bool showRestartAction = false)
    {
        LoadingPage.Visibility = Visibility.Hidden;
        ResultPage.Visibility = Visibility.Visible;
        ResultActionsPanel.Visibility = Visibility.Visible;
        OpenQuickInputButton.Visibility = showQuickInputAction ? Visibility.Visible : Visibility.Collapsed;
        RestartAppButton.Visibility = showRestartAction ? Visibility.Visible : Visibility.Collapsed;
        OpenLogFileButton.Visibility = Visibility.Visible;

        ResultStateIcon.Symbol = SymbolRegular.ErrorCircle24;
        ResultIconBadge.Background = _errorBadgeBrush;
        ResultStateIcon.Foreground = new SolidColorBrush(Colors.White);

        ResultTitleTextBlock.Text = "Dịch thất bại";
        ResultMessageTextBlock.Text = NormalizeSingleLine(subtitle, 96);
        TitleTextBlock.Text = "Có lỗi xảy ra";
        EnsureCurrentPositionIsVisible();
    }

    public void ScheduleClose(TimeSpan delay)
    {
        // Ghi nhớ delay gốc (3s success / 10s error) để khi hover-out sẽ đếm lại đúng từ đầu.
        _scheduledCloseDelay = delay < TimeSpan.Zero ? TimeSpan.Zero : delay;
        _autoCloseRequested = true;
        if (_isPointerHovering)
        {
            CancelScheduledClose();
            return;
        }

        StartCloseCountdown(_scheduledCloseDelay);
    }

    private void ShellBorder_OnMouseEnter(object sender, MouseEventArgs e)
    {
        _isPointerHovering = true;
        if (_autoCloseRequested)
        {
            CancelScheduledClose();
        }
    }

    private void ShellBorder_OnMouseLeave(object sender, MouseEventArgs e)
    {
        _isPointerHovering = false;
        if (_autoCloseRequested)
        {
            // Hover xong rời chuột: reset countdown từ đầu theo delay đã lên lịch.
            StartCloseCountdown(_scheduledCloseDelay);
        }
    }

    private void ShellBorder_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed)
        {
            return;
        }

        if (e.OriginalSource is DependencyObject source &&
            FindVisualParent<System.Windows.Controls.Primitives.ButtonBase>(source) is not null)
        {
            return;
        }

        try
        {
            DragMove();
            SaveCurrentPositionAsCustom();
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopup.ShellBorder_OnMouseLeftButtonDown", ex);
        }
    }

    private void StartCloseCountdown(TimeSpan delay)
    {
        CancelScheduledClose();
        _closeCts = new CancellationTokenSource();
        var token = _closeCts.Token;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(delay, token);
                if (token.IsCancellationRequested)
                {
                    return;
                }

                await Dispatcher.InvokeAsync(() =>
                {
                    if (_isPointerHovering)
                    {
                        return;
                    }

                    CloseSafely();
                });
            }
            catch (OperationCanceledException)
            {
                // Ignore cancel.
            }
        }, token);
    }

    private void PositionAtPreferredLocation()
    {
        // Chỉ dùng lại vị trí đã lưu nếu popup còn nằm gọn trên một màn hình thực tế.
        if (_positionStore.TryLoad(out var savedPosition) && IsPositionVisibleOnAnyScreen(savedPosition))
        {
            Left = savedPosition.X;
            Top = savedPosition.Y;
            _isUsingCustomPosition = true;
            return;
        }

        if (_positionStore.TryLoad(out _))
        {
            _positionStore.Clear();
        }

        ApplyBottomRightPosition();
        _isUsingCustomPosition = false;
    }

    private void ApplyBottomRightPosition()
    {
        if (!TryGetPopupSize(out var width, out var height))
        {
            return;
        }

        var workArea = ResolveCurrentMonitorWorkArea();
        var defaultPosition = BuildBottomRightPosition(workArea, width, height);
        Left = defaultPosition.X;
        Top = defaultPosition.Y;
    }

    private void EnsureCurrentPositionIsVisible()
    {
        if (!_isUsingCustomPosition)
        {
            ApplyBottomRightPosition();
            return;
        }

        var currentPosition = new Point(Left, Top);
        if (IsPositionVisibleOnAnyScreen(currentPosition))
        {
            return;
        }

        // Layout hoặc cấu hình màn hình có thể đã đổi sau lần kéo cuối; reset về mặc định an toàn.
        _positionStore.Clear();
        _isUsingCustomPosition = false;
        ApplyBottomRightPosition();
    }

    private void SaveCurrentPositionAsCustom()
    {
        var currentPosition = new Point(Left, Top);
        if (!IsPositionVisibleOnAnyScreen(currentPosition))
        {
            _positionStore.Clear();
            _isUsingCustomPosition = false;
            ApplyBottomRightPosition();
            return;
        }

        _positionStore.Save(currentPosition);
        _isUsingCustomPosition = true;
    }

    private bool IsPositionVisibleOnAnyScreen(Point position)
    {
        if (!TryGetPopupSize(out var width, out var height))
        {
            return false;
        }

        var popupBounds = new Rect(position.X, position.Y, width, height);
        foreach (var workArea in ResolveAllMonitorWorkAreas())
        {
            if (workArea.Contains(popupBounds.TopLeft) && workArea.Contains(popupBounds.BottomRight))
            {
                return true;
            }
        }

        return false;
    }

    private IEnumerable<Rect> ResolveAllMonitorWorkAreas()
    {
        var workAreas = new List<Rect>();
        var source = PresentationSource.FromVisual(this);
        var transformFromDevice = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;

        // Duyệt từng monitor thật thay vì chỉ dùng virtual screen để tránh coi khoảng trống giữa màn hình là hợp lệ.
        var callback = new NativeMonitorEnumProc((monitor, _, _, _) =>
        {
            var monitorInfo = new NativeMonitorInfo
            {
                Size = Marshal.SizeOf<NativeMonitorInfo>()
            };

            if (!GetMonitorInfo(monitor, ref monitorInfo))
            {
                return true;
            }

            var topLeft = transformFromDevice.Transform(new Point(monitorInfo.Work.Left, monitorInfo.Work.Top));
            var bottomRight = transformFromDevice.Transform(new Point(monitorInfo.Work.Right, monitorInfo.Work.Bottom));
            workAreas.Add(new Rect(topLeft, bottomRight));
            return true;
        });

        if (!EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, callback, IntPtr.Zero) || workAreas.Count == 0)
        {
            workAreas.Add(SystemParameters.WorkArea);
        }

        return workAreas;
    }

    private Point BuildBottomRightPosition(Rect workArea, double width, double height)
    {
        var targetLeft = workArea.Right - width - EdgeMargin;
        var targetTop = workArea.Bottom - height - EdgeMargin;

        var minLeft = workArea.Left + EdgeMargin;
        var minTop = workArea.Top + EdgeMargin;

        var maxLeft = workArea.Right - width - EdgeMargin;
        var maxTop = workArea.Bottom - height - EdgeMargin;

        if (maxLeft < minLeft)
        {
            maxLeft = minLeft;
        }

        if (maxTop < minTop)
        {
            maxTop = minTop;
        }

        return new Point(
            Math.Clamp(targetLeft, minLeft, maxLeft),
            Math.Clamp(targetTop, minTop, maxTop)
        );
    }

    private bool TryGetPopupSize(out double width, out double height)
    {
        width = ActualWidth;
        if (width <= 0 || double.IsNaN(width))
        {
            width = Width;
        }

        height = ActualHeight;
        if (height <= 0 || double.IsNaN(height))
        {
            height = Height;
        }

        if (double.IsNaN(width) || double.IsNaN(height) || width <= 0 || height <= 0)
        {
            UpdateLayout();
            width = ActualWidth;
            height = ActualHeight;
        }

        if (double.IsNaN(width) || double.IsNaN(height) || width <= 0 || height <= 0)
        {
            width = 400;
            height = 120;
        }

        return width > 0 && height > 0;
    }

    private Rect ResolveCurrentMonitorWorkArea()
    {
        if (!GetCursorPos(out var cursorPoint))
        {
            return SystemParameters.WorkArea;
        }

        var monitor = MonitorFromPoint(cursorPoint, MonitorDefaultToNearest);
        if (monitor == IntPtr.Zero)
        {
            return SystemParameters.WorkArea;
        }

        var monitorInfo = new NativeMonitorInfo
        {
            Size = Marshal.SizeOf<NativeMonitorInfo>()
        };
        if (!GetMonitorInfo(monitor, ref monitorInfo))
        {
            return SystemParameters.WorkArea;
        }

        var source = PresentationSource.FromVisual(this);
        var transformFromDevice = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;

        var topLeft = transformFromDevice.Transform(new Point(monitorInfo.Work.Left, monitorInfo.Work.Top));
        var bottomRight = transformFromDevice.Transform(new Point(monitorInfo.Work.Right, monitorInfo.Work.Bottom));
        return new Rect(topLeft, bottomRight);
    }

    private static T? FindVisualParent<T>(DependencyObject source) where T : DependencyObject
    {
        var current = source;
        while (current is not null)
        {
            if (current is T match)
            {
                return match;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private void CancelScheduledClose()
    {
        if (_closeCts is null)
        {
            return;
        }

        _closeCts.Cancel();
        _closeCts.Dispose();
        _closeCts = null;
    }

    protected override void OnClosed(EventArgs e)
    {
        CancelScheduledClose();
        _autoCloseRequested = false;
        base.OnClosed(e);
    }

    private void CloseButton_OnClick(object sender, RoutedEventArgs e)
    {
        _autoCloseRequested = false;
        CloseSafely();
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape)
        {
            return;
        }

        e.Handled = true;
        _autoCloseRequested = false;
        CloseSafely();
    }

    private void OpenQuickInputButton_OnClick(object sender, RoutedEventArgs e)
    {
        _autoCloseRequested = false;
        QuickInputRequested?.Invoke(this, EventArgs.Empty);
    }

    private void RestartAppButton_OnClick(object sender, RoutedEventArgs e)
    {
        _autoCloseRequested = false;
        RestartAppRequested?.Invoke(this, EventArgs.Empty);
    }

    private void OpenLogFileButton_OnClick(object sender, RoutedEventArgs e)
    {
        _autoCloseRequested = false;

        try
        {
            var logPath = ErrorFileLogger.GetTextLogPath();
            if (!File.Exists(logPath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
                File.WriteAllText(logPath, string.Empty);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = logPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopup.OpenLogFileButton_OnClick", ex);
            System.Windows.MessageBox.Show(
                $"Không mở được log file: {ex.Message}",
                "Instant Translate",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning
            );
        }
    }

    private static string NormalizeSingleLine(string? text, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return "chưa có nội dung";
        }

        var normalized = text.Replace('\r', ' ').Replace('\n', ' ').Trim();
        while (normalized.Contains("  ", StringComparison.Ordinal))
        {
            normalized = normalized.Replace("  ", " ", StringComparison.Ordinal);
        }

        if (normalized.Length <= maxLength)
        {
            return normalized;
        }

        return $"{normalized[..(maxLength - 3)]}...";
    }

    private void SetCurrentStep(string step, SymbolRegular icon)
    {
        StepStatusTextBlock.Text = step;
        StepStatusIcon.Symbol = icon;
        StepIconBadge.Background = GetStepBadgeBrush(icon);
    }

    private SymbolRegular ResolveStepIcon(string stepText)
    {
        var normalized = stepText.ToLowerInvariant();

        if (normalized.Contains("lỗi", StringComparison.Ordinal) || normalized.Contains("thất bại", StringComparison.Ordinal))
        {
            return SymbolRegular.ErrorCircle24;
        }

        if (normalized.Contains("clipboard", StringComparison.Ordinal) || normalized.Contains("copy", StringComparison.Ordinal))
        {
            return SymbolRegular.Clipboard24;
        }

        if (normalized.Contains("lưu", StringComparison.Ordinal) || normalized.Contains("lịch sử", StringComparison.Ordinal))
        {
            return SymbolRegular.History24;
        }

        if (normalized.Contains("kết quả", StringComparison.Ordinal)
            || normalized.Contains("hoàn tất", StringComparison.Ordinal)
            || normalized.Contains("thành công", StringComparison.Ordinal))
        {
            return SymbolRegular.Checkmark24;
        }

        if (normalized.Contains("dịch", StringComparison.Ordinal))
        {
            return SymbolRegular.Translate24;
        }

        if (normalized.Contains("khởi tạo", StringComparison.Ordinal)
            || normalized.Contains("nhận tác vụ", StringComparison.Ordinal)
            || normalized.Contains("chuẩn bị", StringComparison.Ordinal))
        {
            return SymbolRegular.Rocket24;
        }

        return SymbolRegular.ArrowSync24;
    }

    private Brush GetStepBadgeBrush(SymbolRegular icon)
    {
        return icon switch
        {
            SymbolRegular.Checkmark24 => _successBadgeBrush,
            SymbolRegular.ErrorCircle24 => _errorBadgeBrush,
            SymbolRegular.Clipboard24 => _clipboardBadgeBrush,
            SymbolRegular.History24 => _historyBadgeBrush,
            SymbolRegular.Translate24 => _processingBadgeBrush,
            SymbolRegular.ArrowSync24 => _processingBadgeBrush,
            _ => _infoBadgeBrush
        };
    }

    private void CloseSafely()
    {
        try
        {
            Close();
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopup.CloseSafely", ex);
            // Ignore close failures for non-critical popup.
        }
    }
}
