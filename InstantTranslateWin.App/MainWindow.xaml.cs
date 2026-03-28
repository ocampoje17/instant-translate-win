using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using InstantTranslateWin.App.Models;
using InstantTranslateWin.App.Services;
using Wpf.Ui.Appearance;
using Wpf.Ui;
using Wpf.Ui.Controls;
using Wpf.Ui.Tray.Controls;

namespace InstantTranslateWin.App;

public partial class MainWindow : FluentWindow
{
    private const int HotkeyId = 1001;
    private const string DefaultTargetLanguage = "English";
    private const string ApiProviderGemini = "Gemini";
    private const string ApiProviderLocalAi = "LocalAi";
    private const string DefaultLocalAiBaseUrl = "https://api.openai.com/v1";
    private const string LegacyDefaultLocalAiBaseUrl = "http://localhost:1234/v1";
    private const string DefaultLocalAiModel = "gpt-4o-mini";
    private static readonly string[] DefaultModels = ["gemini-flash-lite-latest", "gemini-2.0-flash-lite"];
    private static readonly string[] HotkeyKeys = BuildHotkeyKeys();
    private static readonly LanguageOption[] SupportedLanguages =
    [
        new() { DisplayName = "English", PromptName = "English" },
        new() { DisplayName = "Vietnamese (Tiếng Việt)", PromptName = "Vietnamese" },
        new() { DisplayName = "Japanese (日本語)", PromptName = "Japanese" },
        new() { DisplayName = "Korean (한국어)", PromptName = "Korean" },
        new() { DisplayName = "Chinese Simplified (简体中文)", PromptName = "Simplified Chinese" },
        new() { DisplayName = "Chinese Traditional (繁體中文)", PromptName = "Traditional Chinese" },
        new() { DisplayName = "French", PromptName = "French" },
        new() { DisplayName = "German", PromptName = "German" },
        new() { DisplayName = "Spanish", PromptName = "Spanish" },
        new() { DisplayName = "Portuguese", PromptName = "Portuguese" },
        new() { DisplayName = "Thai (ไทย)", PromptName = "Thai" },
        new() { DisplayName = "Indonesian", PromptName = "Indonesian" }
    ];

    private readonly AppStateStore _stateStore = new();
    private readonly GeminiTranslationService _translationService = new();
    private readonly LocalAiTranslationService _localAiTranslationService = new();
    private readonly SelectedTextCaptureService _captureService = new();
    private readonly StartupRegistrationService _startupRegistrationService = new();
    private readonly ISnackbarService _snackbarService = new SnackbarService();
    private readonly SemaphoreSlim _translateLock = new(1, 1);
    private readonly SemaphoreSlim _stateSaveLock = new(1, 1);
    private readonly Queue<PendingTranslationRequest> _translationQueue = new();
    private readonly object _translationQueueLock = new();
    private readonly object _manualTranslationLock = new();

    private AppState _state = new();
    private GlobalHotkeyService? _hotkeyService;
    private HotkeyProgressPopup? _hotkeyPopup;
    private readonly Stack<NavigationViewItem> _navigationBackStack = new();
    private string? _currentTranslatingInput;
    private string? _currentManualTranslatingInput;
    private CancellationTokenSource? _manualTranslationCts;
    private Task _manualTranslationTask = Task.CompletedTask;
    private int _manualTranslationVersion;
    private bool _allowClose;
    private bool _isHotkeyRegistered;
    private bool _isShuttingDown;
    private bool _isApplyingSettings;
    private bool _isAutoSavingSettings;
    private bool _isManualTargetLanguageInitialized;

    public ObservableCollection<TranslationRecord> HistoryRecords { get; } = [];
    public ObservableCollection<ApiKeyInputItem> ApiKeyInputItems { get; } = [];
    public ObservableCollection<string> HotkeyKeyOptions { get; } = [];
    public ObservableCollection<LanguageOption> TargetLanguageOptions { get; } = [];
    public ObservableCollection<ThemeOptionItem> ThemeOptions { get; } = [];

    private sealed record PendingTranslationRequest(
        string SourceText,
        AppSettings Settings,
        IReadOnlyList<string> ApiKeys
    );

    private enum ProviderValidationErrorKind
    {
        None,
        MissingApiKey,
        InvalidConfiguration
    }

    public sealed record ThemeOptionItem(string DisplayName, string ThemeKey);

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;

        Loaded += OnLoaded;
        Closing += OnClosing;
        Closed += OnClosed;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        _snackbarService.SetSnackbarPresenter(StatusSnackbarPresenter);
        _snackbarService.DefaultTimeOut = TimeSpan.FromSeconds(3);

        InitializeHotkeyOptions();
        InitializeLanguageOptions();
        InitializeThemeOptions();

        _state = await _stateStore.LoadAsync();
        var migrated = EnsureSettingsCompatibility(_state.Settings);
        try
        {
            var startupEnabled = _startupRegistrationService.IsEnabled();
            if (_state.Settings.LaunchOnStartup != startupEnabled)
            {
                _state.Settings.LaunchOnStartup = startupEnabled;
                migrated = true;
            }
        }
        catch
        {
            // Ignore startup key read issues; UI will keep persisted state.
        }

        ApplySettingsToUi(_state.Settings);
        ApplyTheme(_state.Settings.AppTheme);

        HistoryRecords.Clear();
        foreach (var record in _state.History.OrderByDescending(x => x.Timestamp))
        {
            HistoryRecords.Add(record);
        }

        AppNotifyIcon.Register();

        _isHotkeyRegistered = TryRegisterHotkey(_state.Settings, out var error);
        if (!_isHotkeyRegistered)
        {
            ShowStatus("Không đăng ký được phím tắt", error, InfoBarSeverity.Warning);
        }
        else
        {
            ShowStatus(
                "Sẵn sàng",
                $"Hotkey: {BuildHotkeyDisplay(_state.Settings)} | Ngôn ngữ đích: {GetTargetLanguageDisplayName(_state.Settings.TargetLanguage)}",
                InfoBarSeverity.Success
            );
        }

        if (migrated)
        {
            await PersistStateAsync();
        }

        SetMainTranslationIdleStatus();
        SetManualTranslationIdleStatus();
        _navigationBackStack.Clear();
        SwitchToTab(0, QuickNavItem);
    }

    private void OnClosing(object? sender, CancelEventArgs e)
    {
        if (_isShuttingDown)
        {
            return;
        }

        if (_allowClose)
        {
            _ = PersistSettingsOnCloseAsync();
            return;
        }

        var keepRunningInBackground = KeepRunningInBackgroundOnCloseCheckBox.IsChecked == true;
        if (keepRunningInBackground)
        {
            e.Cancel = true;
            Hide();
            ShowInTaskbar = false;
            _ = PersistSettingsOnCloseAsync();
            return;
        }

        _ = PersistSettingsOnCloseAsync();
        AppNotifyIcon.Unregister();
    }

    private async void OnClosed(object? sender, EventArgs e)
    {
        _isShuttingDown = true;

        Task manualTaskToAwait;
        lock (_manualTranslationLock)
        {
            _manualTranslationVersion++;
            _manualTranslationCts?.Cancel();
            manualTaskToAwait = _manualTranslationTask;
            _currentManualTranslatingInput = null;
        }

        lock (_translationQueueLock)
        {
            _translationQueue.Clear();
            _currentTranslatingInput = null;
        }

        await _translateLock.WaitAsync();
        try
        {
            _hotkeyService?.Dispose();
            if (_hotkeyPopup is not null)
            {
                try
                {
                    _hotkeyPopup.Close();
                }
                catch
                {
                    // Ignore popup close failures during app shutdown.
                }
            }

            AppNotifyIcon.Dispose();
            _translationService.Dispose();
            _localAiTranslationService.Dispose();
            try
            {
                await manualTaskToAwait;
            }
            catch (OperationCanceledException)
            {
                // Ignore cancellation during app shutdown.
            }
            catch
            {
                // Ignore background failures during app shutdown.
            }

            await PersistStateAsync();
        }
        finally
        {
            _translateLock.Release();
        }
    }

    private void AppNotifyIcon_OnLeftDoubleClick(NotifyIcon sender, RoutedEventArgs e)
    {
        RestoreWindowFromTray();
    }

    private void RootNavigationView_SelectionChanged(NavigationView sender, RoutedEventArgs args)
    {
        if (sender.SelectedItem is not INavigationViewItem item)
        {
            return;
        }

        switch (item.TargetPageTag)
        {
            case "quick":
                SwitchToTab(0, QuickNavItem);
                break;
            case "manual":
                SwitchToTab(1, ManualNavItem);
                break;
            case "history":
                SwitchToTab(2, HistoryNavItem);
                break;
            case "settings":
                SwitchToTab(3, SettingsNavItem);
                break;
        }
    }

    private void QuickNavItem_Click(object sender, RoutedEventArgs e)
    {
        SwitchToTab(0, QuickNavItem);
    }

    private void ManualNavItem_Click(object sender, RoutedEventArgs e)
    {
        SwitchToTab(1, ManualNavItem);
    }

    private void HistoryNavItem_Click(object sender, RoutedEventArgs e)
    {
        SwitchToTab(2, HistoryNavItem);
    }

    private void SettingsNavItem_Click(object sender, RoutedEventArgs e)
    {
        SwitchToTab(3, SettingsNavItem);
    }

    private void RootNavigationView_BackRequested(NavigationView sender, RoutedEventArgs args)
    {
        NavigateBack();
    }

    private void NavigateBack()
    {
        if (_navigationBackStack.Count == 0)
        {
            return;
        }

        var previousItem = _navigationBackStack.Pop();
        var index = previousItem.TargetPageTag switch
        {
            "quick" => 0,
            "manual" => 1,
            "history" => 2,
            "settings" => 3,
            _ => -1
        };

        if (index >= 0)
        {
            SwitchToTab(index, previousItem, isBackNavigation: true);
        }
        else
        {
            UpdateBackNavigationState();
        }
    }

    private void SwitchToTab(int index, NavigationViewItem activeItem, bool isBackNavigation = false)
    {
        var previousItem = ContentTabControl.SelectedIndex switch
        {
            0 => QuickNavItem,
            1 => ManualNavItem,
            2 => HistoryNavItem,
            3 => SettingsNavItem,
            _ => null
        };

        if (!isBackNavigation && previousItem is not null && !ReferenceEquals(previousItem, activeItem))
        {
            _navigationBackStack.Push(previousItem);
        }

        ContentTabControl.SelectedIndex = index;
        activeItem.Activate(RootNavigationView);

        QuickNavItem.IsActive = ReferenceEquals(activeItem, QuickNavItem);
        ManualNavItem.IsActive = ReferenceEquals(activeItem, ManualNavItem);
        HistoryNavItem.IsActive = ReferenceEquals(activeItem, HistoryNavItem);
        SettingsNavItem.IsActive = ReferenceEquals(activeItem, SettingsNavItem);
        UpdateBackNavigationState();
    }

    private void UpdateBackNavigationState()
    {
        RootNavigationView.SetCurrentValue(NavigationView.IsBackEnabledProperty, _navigationBackStack.Count > 0);
    }

    private void TrayOpenMenuItem_Click(object sender, RoutedEventArgs e)
    {
        RestoreWindowFromTray();
    }

    private void TrayExitMenuItem_Click(object sender, RoutedEventArgs e)
    {
        _allowClose = true;
        AppNotifyIcon.Unregister();
        Close();
    }

    private void RestoreWindowFromTray()
    {
        ShowInTaskbar = true;
        Show();
        WindowState = WindowState.Normal;
        Activate();
    }

    private HotkeyProgressPopup? ShowHotkeyProgressPopup()
    {
        if (_hotkeyPopup is null)
        {
            _hotkeyPopup = new HotkeyProgressPopup();
            _hotkeyPopup.RestartAppRequested += HotkeyPopupOnRestartAppRequested;
        }

        if (!TryShowHotkeyPopup(_hotkeyPopup))
        {
            try
            {
                _hotkeyPopup.Close();
            }
            catch
            {
                // Ignore close failure and recreate popup below.
            }

            _hotkeyPopup = new HotkeyProgressPopup();
            _hotkeyPopup.RestartAppRequested += HotkeyPopupOnRestartAppRequested;
            if (!TryShowHotkeyPopup(_hotkeyPopup))
            {
                _hotkeyPopup = null;
                ShowStatus("Không hiển thị được popup", "Popup trạng thái tạm thời không khả dụng.", InfoBarSeverity.Warning);
                return null;
            }
        }

        _hotkeyPopup.BeginProgress("Đang xử lý hotkey...", string.Empty);
        return _hotkeyPopup;
    }

    private void HotkeyPopupOnRestartAppRequested(object? sender, EventArgs e)
    {
        try
        {
            var currentExePath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(currentExePath))
            {
                ShowStatus("Không thể khởi động lại app", "Không xác định được đường dẫn ứng dụng hiện tại.", InfoBarSeverity.Error);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = currentExePath,
                UseShellExecute = true,
                WorkingDirectory = AppContext.BaseDirectory
            });

            _allowClose = true;
            AppNotifyIcon.Unregister();
            Close();
        }
        catch (Exception ex)
        {
            ShowStatus("Không thể khởi động lại app", ex.Message, InfoBarSeverity.Error);
        }
    }

    private static bool TryShowHotkeyPopup(HotkeyProgressPopup popup)
    {
        try
        {
            popup.ShowAtBottomRight();
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void BeginMainTranslationProgress(string status, double progressPercent = 0)
    {
        UpdateMainTranslationProgress(status, progressPercent);
    }

    private void UpdateMainTranslationProgress(string status, double progressPercent)
    {
        QuickTranslationProgressView.SetProgress(status, progressPercent);
    }

    private void SetMainTranslationIdleStatus()
    {
        var status = _isHotkeyRegistered
            ? $"Sẵn sàng. Hotkey: {BuildHotkeyDisplay(_state.Settings)}"
            : "Hotkey chưa sẵn sàng. Mở tab Cài đặt để cấu hình lại.";
        QuickTranslationProgressView.SetIdle(status);
    }

    private void BeginManualTranslationProgress(string status, double progressPercent = 0)
    {
        UpdateManualTranslationProgress(status, progressPercent);
    }

    private void UpdateManualTranslationProgress(string status, double progressPercent)
    {
        ManualTranslationProgressView.SetProgress(status, progressPercent);
    }

    private void SetManualTranslationIdleStatus()
    {
        ManualTranslationProgressView.SetIdle(
            $"Sẵn sàng. Ngôn ngữ đích: {GetTargetLanguageDisplayName(GetManualTargetLanguage())}"
        );
    }

    private async void HotkeyServiceOnHotkeyPressed(object? sender, EventArgs e)
    {
        await TranslateSelectedTextFromForegroundAsync();
    }

    private async Task TranslateSelectedTextFromForegroundAsync()
    {
        try
        {
            var settings = BuildSettingsFromUi(includeApiKey: false);
            if (
                !TryResolveProviderRequest(
                    settings,
                    out var apiKeys,
                    out var validationError,
                    out var validationErrorKind
                )
            )
            {
                ShowProviderValidationError(validationError, validationErrorKind);
                return;
            }

            var selectedText = await _captureService.CaptureSelectedTextAsync();

            if (string.IsNullOrWhiteSpace(selectedText))
            {
                var popup = ShowHotkeyProgressPopup();
                if (popup is not null)
                {
                    popup.AddStep("Không thể đọc text đang chọn.");
                    popup.SetErrorState("Hãy chọn lại đoạn văn bản rồi thử lại. Nếu vẫn lỗi, hãy khởi động lại app.", showRestartAction: true);
                }

                ShowStatus(
                    "Không có text",
                    "Không đọc được text đang chọn. Hãy thử chọn lại; nếu vẫn lỗi, hãy khởi động lại app.",
                    InfoBarSeverity.Warning
                );
                return;
            }

            var request = new PendingTranslationRequest(selectedText, settings, apiKeys);
            await EnqueueOrProcessTranslationAsync(request);
        }
        catch (Exception ex)
        {
            ShowStatus("Lỗi dịch", ex.Message, InfoBarSeverity.Error);
        }
    }

    private void TranslateManualButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var input = ManualSourceTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                ShowStatus("Thiếu nội dung", "Nhập text cần dịch ở tab Dịch thủ công.", InfoBarSeverity.Warning);
                return;
            }

            var settings = BuildSettingsFromUi(includeApiKey: false);
            settings.TargetLanguage = GetManualTargetLanguage();
            if (
                !TryResolveProviderRequest(
                    settings,
                    out var apiKeys,
                    out var validationError,
                    out var validationErrorKind
                )
            )
            {
                ShowProviderValidationError(validationError, validationErrorKind);
                return;
            }

            StartOrReplaceManualTranslation(input, settings, apiKeys);
        }
        catch (Exception ex)
        {
            ShowStatus("Lỗi dịch", ex.Message, InfoBarSeverity.Error);
        }
    }

    private void ManualTargetLanguageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!HasActiveManualTranslation())
        {
            SetManualTranslationIdleStatus();
        }
    }

    private void StartOrReplaceManualTranslation(string sourceText, AppSettings settings, IReadOnlyList<string> apiKeys)
    {
        if (_isShuttingDown)
        {
            return;
        }

        var normalizedInput = NormalizeQueueInput(sourceText);
        var ignoreAsDuplicate = false;
        var canceledPrevious = false;
        var requestVersion = 0;
        CancellationTokenSource? requestCts = null;

        lock (_manualTranslationLock)
        {
            if (_currentManualTranslatingInput is not null && _manualTranslationCts is not null)
            {
                if (string.Equals(_currentManualTranslatingInput, normalizedInput, StringComparison.Ordinal))
                {
                    ignoreAsDuplicate = true;
                }
                else
                {
                    _manualTranslationCts.Cancel();
                    canceledPrevious = true;
                }
            }

            if (!ignoreAsDuplicate)
            {
                requestCts = new CancellationTokenSource();
                _manualTranslationCts = requestCts;
                _currentManualTranslatingInput = normalizedInput;
                requestVersion = ++_manualTranslationVersion;
            }
        }

        if (ignoreAsDuplicate)
        {
            ShowStatus("Bỏ qua yêu cầu trùng", "Nội dung này đang được dịch thủ công.", InfoBarSeverity.Informational);
            return;
        }

        if (requestCts is null)
        {
            return;
        }

        if (canceledPrevious)
        {
            ShowStatus("Đã hủy yêu cầu cũ", "Đang chạy yêu cầu dịch thủ công mới nhất.", InfoBarSeverity.Informational);
        }

        var manualTask = ProcessManualTranslationAsync(sourceText, settings, apiKeys, requestVersion, requestCts);
        lock (_manualTranslationLock)
        {
            if (_manualTranslationVersion == requestVersion)
            {
                _manualTranslationTask = manualTask;
            }
        }
    }

    private async Task ProcessManualTranslationAsync(
        string sourceText,
        AppSettings settings,
        IReadOnlyList<string> apiKeys,
        int requestVersion,
        CancellationTokenSource requestCts
    )
    {
        var cancellationToken = requestCts.Token;

        try
        {
            BeginManualTranslationProgress("Đang chuẩn bị dịch thủ công...", 10);
            UpdateManualTranslationProgress($"Đang dịch với {GetProviderDisplayName(settings.ActiveApiProvider)}...", 55);

            var translation = await TranslateWithSettingsAsync(sourceText, settings, apiKeys, cancellationToken);
            if (cancellationToken.IsCancellationRequested || !IsCurrentManualRequest(requestVersion))
            {
                return;
            }

            UpdateManualTranslationProgress("Đã nhận kết quả dịch.", 80);
            ManualResultTextBox.Text = translation;

            AddToHistory(sourceText, translation, GetActiveModelName(settings));
            await PersistStateAsync();
            if (cancellationToken.IsCancellationRequested || !IsCurrentManualRequest(requestVersion))
            {
                return;
            }

            UpdateManualTranslationProgress("Dịch thủ công thành công.", 100);
            ShowStatus("Dịch thủ công thành công", "Kết quả đã thêm vào lịch sử.", InfoBarSeverity.Success);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // Replaced by a newer manual request.
        }
        catch (Exception ex)
        {
            if (cancellationToken.IsCancellationRequested || !IsCurrentManualRequest(requestVersion))
            {
                return;
            }

            UpdateManualTranslationProgress("Xảy ra lỗi trong quá trình dịch thủ công.", 100);
            ShowStatus("Lỗi dịch thủ công", ex.Message, InfoBarSeverity.Error);
        }
        finally
        {
            requestCts.Dispose();

            var shouldResetManualStatus = false;
            lock (_manualTranslationLock)
            {
                if (_manualTranslationVersion == requestVersion)
                {
                    _manualTranslationCts = null;
                    _currentManualTranslatingInput = null;
                    shouldResetManualStatus = true;
                }
            }

            if (shouldResetManualStatus && !_isShuttingDown)
            {
                SetManualTranslationIdleStatus();
            }
        }
    }

    private bool IsCurrentManualRequest(int requestVersion)
    {
        lock (_manualTranslationLock)
        {
            return _manualTranslationVersion == requestVersion;
        }
    }

    private async Task EnqueueOrProcessTranslationAsync(PendingTranslationRequest request)
    {
        if (_isShuttingDown)
        {
            return;
        }

        var normalizedInput = NormalizeQueueInput(request.SourceText);
        var shouldStartNow = false;
        var queued = false;
        var duplicateCurrent = false;
        var queuedCount = 0;

        lock (_translationQueueLock)
        {
            if (_currentTranslatingInput is null)
            {
                _currentTranslatingInput = normalizedInput;
                shouldStartNow = true;
            }
            else if (string.Equals(_currentTranslatingInput, normalizedInput, StringComparison.Ordinal))
            {
                duplicateCurrent = true;
            }
            else
            {
                _translationQueue.Enqueue(request);
                queued = true;
                queuedCount = _translationQueue.Count;
            }
        }

        if (duplicateCurrent)
        {
            ShowStatus("Bỏ qua yêu cầu trùng", "Nội dung này đang được dịch, không tạo tác vụ mới.", InfoBarSeverity.Informational);
            return;
        }

        if (queued)
        {
            ShowStatus("Đã thêm vào hàng chờ", $"Số tác vụ đang chờ: {queuedCount}.", InfoBarSeverity.Informational);
            return;
        }

        if (shouldStartNow)
        {
            await ProcessTranslationQueueAsync(request);
        }
    }

    private async Task ProcessTranslationQueueAsync(PendingTranslationRequest initialRequest)
    {
        var request = initialRequest;

        while (true)
        {
            if (_isShuttingDown)
            {
                return;
            }

            await ProcessSingleTranslationRequestAsync(request);

            PendingTranslationRequest? nextRequest = null;
            var remainingQueueCount = 0;

            lock (_translationQueueLock)
            {
                if (_isShuttingDown)
                {
                    _translationQueue.Clear();
                    _currentTranslatingInput = null;
                    return;
                }

                if (_translationQueue.Count > 0)
                {
                    nextRequest = _translationQueue.Dequeue();
                    _currentTranslatingInput = NormalizeQueueInput(nextRequest.SourceText);
                    remainingQueueCount = _translationQueue.Count;
                }
                else
                {
                    _currentTranslatingInput = null;
                }
            }

            if (nextRequest is null)
            {
                SetMainTranslationIdleStatus();
                return;
            }

            ShowStatus("Đang xử lý hàng chờ", $"Còn {remainingQueueCount} tác vụ trong hàng chờ.", InfoBarSeverity.Informational);
            request = nextRequest;
        }
    }

    private async Task ProcessSingleTranslationRequestAsync(PendingTranslationRequest request)
    {
        await _translateLock.WaitAsync();
        HotkeyProgressPopup? popup = null;

        try
        {
            BeginMainTranslationProgress("Đang chuẩn bị dịch từ hotkey...", 10);
            popup = ShowHotkeyProgressPopup();
            popup?.AddStep("Đã nhận tác vụ dịch.");
            popup?.SetSourcePreview(request.SourceText);
            popup?.SetProgressState($"Đang dịch với {GetProviderDisplayName(request.Settings.ActiveApiProvider)}...", 55);

            UpdateMainTranslationProgress($"Đang dịch với {GetProviderDisplayName(request.Settings.ActiveApiProvider)}...", 55);
            var translation = await TranslateWithSettingsAsync(request.SourceText, request.Settings, request.ApiKeys);
            UpdateMainTranslationProgress("Đã nhận kết quả dịch.", 80);

            popup?.AddStep("Đã nhận kết quả dịch.");
            popup?.SetProgressState("Đã nhận kết quả dịch.", 80);
            LastSourceTextBox.Text = request.SourceText;
            LastTranslatedTextBox.Text = translation;

            if (request.Settings.CopyTranslationToClipboard)
            {
                System.Windows.Clipboard.SetText(translation);
                popup?.AddStep("Đã copy kết quả vào clipboard.");
                popup?.SetProgressState("Đã copy kết quả.", 90);
                UpdateMainTranslationProgress("Đã copy kết quả vào clipboard.", 90);
            }

            AddToHistory(request.SourceText, translation, GetActiveModelName(request.Settings));
            await PersistStateAsync();
            UpdateMainTranslationProgress("Đã lưu lịch sử. Đang hoàn tất...", 96);

            popup?.AddStep("Đã lưu lịch sử.");
            popup?.SetProgressState("Đang hoàn tất...", 96);
            popup?.SetSuccessState(
                $"Ngôn ngữ đích: {GetTargetLanguageDisplayName(request.Settings.TargetLanguage)}",
                translation
            );
            popup?.ScheduleClose(TimeSpan.FromSeconds(3));

            ShowStatus(
                "Dịch thành công",
                request.Settings.CopyTranslationToClipboard
                    ? "Bản dịch đã được copy vào clipboard."
                    : "Đã lưu bản dịch vào lịch sử.",
                InfoBarSeverity.Success
            );

            UpdateMainTranslationProgress("Dịch thành công.", 100);
        }
        catch (Exception ex)
        {
            if (popup is not null)
            {
                popup.AddStep("Xảy ra lỗi trong quá trình dịch.");
                popup.SetErrorState(ex.Message, showRestartAction: false);
                popup.ScheduleClose(TimeSpan.FromSeconds(3));
            }

            UpdateMainTranslationProgress("Xảy ra lỗi trong quá trình dịch.", 100);
            ShowStatus("Lỗi dịch", ex.Message, InfoBarSeverity.Error);
        }
        finally
        {
            _translateLock.Release();
        }
    }

    private async Task PersistSettingsOnCloseAsync()
    {
        if (!IsLoaded || _isApplyingSettings)
        {
            return;
        }

        try
        {
            _state.Settings = BuildSettingsFromUi(includeApiKey: false);
            await PersistStateAsync();
        }
        catch
        {
            // Ignore close-time settings persistence failures.
        }
    }

    private static string NormalizeQueueInput(string? sourceText)
    {
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            return string.Empty;
        }

        return sourceText.Replace("\r\n", "\n", StringComparison.Ordinal).Trim();
    }

    private bool HasActiveTranslation()
    {
        lock (_translationQueueLock)
        {
            return _currentTranslatingInput is not null;
        }
    }

    private bool HasActiveManualTranslation()
    {
        lock (_manualTranslationLock)
        {
            return _currentManualTranslatingInput is not null;
        }
    }

    private string GetManualTargetLanguage()
    {
        var selected = ManualTargetLanguageComboBox.SelectedValue?.ToString();
        if (string.IsNullOrWhiteSpace(selected))
        {
            selected = _state.Settings.TargetLanguage;
        }

        return NormalizeTargetLanguage(selected);
    }

    private List<string> BuildEncryptedApiKeysFromUi()
    {
        var uniquePlainTexts = new HashSet<string>(StringComparer.Ordinal);
        var encryptedKeys = new List<string>();

        foreach (var row in ApiKeyInputItems)
        {
            var plainText = row.PlainText?.Trim();
            if (string.IsNullOrWhiteSpace(plainText))
            {
                continue;
            }

            if (!uniquePlainTexts.Add(plainText))
            {
                continue;
            }

            var encrypted = SecretProtector.Protect(plainText);
            if (!string.IsNullOrWhiteSpace(encrypted))
            {
                encryptedKeys.Add(encrypted);
            }
        }

        return encryptedKeys;
    }

    private static List<string> GetEncryptedApiKeysFromSettings(AppSettings settings)
    {
        var keys = new List<string>();
        if (settings.GeminiApiKeysEncrypted is not null)
        {
            keys.AddRange(settings.GeminiApiKeysEncrypted);
        }

        if (!string.IsNullOrWhiteSpace(settings.GeminiApiKeyEncrypted))
        {
            keys.Add(settings.GeminiApiKeyEncrypted);
        }

        return keys
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }

    private static List<string> GetDecryptedApiKeysFromSettings(AppSettings settings)
    {
        var keys = new List<string>();
        var uniquePlainTexts = new HashSet<string>(StringComparer.Ordinal);

        foreach (var encrypted in GetEncryptedApiKeysFromSettings(settings))
        {
            var plainText = SecretProtector.Unprotect(encrypted).Trim();
            if (string.IsNullOrWhiteSpace(plainText))
            {
                continue;
            }

            if (uniquePlainTexts.Add(plainText))
            {
                keys.Add(plainText);
            }
        }

        return keys;
    }

    private static string? BuildEncryptedLocalAiApiKeyFromPlainText(string? plainText)
    {
        var trimmed = plainText?.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return null;
        }

        return SecretProtector.Protect(trimmed);
    }

    private string? BuildEncryptedLocalAiApiKeyFromUi()
    {
        return BuildEncryptedLocalAiApiKeyFromPlainText(LocalAiApiKeyTextBox.Text);
    }

    private static string? GetDecryptedLocalAiApiKeyFromSettings(AppSettings settings)
    {
        var plainText = SecretProtector.Unprotect(settings.LocalAiApiKeyEncrypted).Trim();
        return string.IsNullOrWhiteSpace(plainText) ? null : plainText;
    }

    private string GetApiProviderFromUi()
    {
        if (ApiProviderLocalAiRadioButton.IsChecked == true)
        {
            return ApiProviderLocalAi;
        }

        if (ApiProviderGeminiRadioButton.IsChecked == true)
        {
            return ApiProviderGemini;
        }

        return NormalizeApiProvider(_state.Settings.ActiveApiProvider);
    }

    private bool TryResolveProviderRequest(
        AppSettings settings,
        out IReadOnlyList<string> apiKeys,
        out string validationError,
        out ProviderValidationErrorKind validationErrorKind
    )
    {
        settings.ActiveApiProvider = NormalizeApiProvider(settings.ActiveApiProvider);
        settings.LocalAiBaseUrl = NormalizeLocalAiBaseUrl(settings.LocalAiBaseUrl);
        settings.LocalAiModelName = NormalizeLocalAiModelName(settings.LocalAiModelName);

        if (IsLocalAiProvider(settings.ActiveApiProvider))
        {
            apiKeys = [];
            var localAiBaseUrl = ResolveLocalAiBaseUrl(settings);
            if (settings.LocalAiUseCustomBaseUrl && !Uri.TryCreate(localAiBaseUrl, UriKind.Absolute, out _))
            {
                validationError = "Base URL tuỳ chỉnh cho OpenAI-compatible không hợp lệ.";
                validationErrorKind = ProviderValidationErrorKind.InvalidConfiguration;
                return false;
            }

            if (!settings.LocalAiUseCustomBaseUrl &&
                string.IsNullOrWhiteSpace(GetDecryptedLocalAiApiKeyFromSettings(settings)))
            {
                validationError = "Bạn chưa nhập API key cho OpenAI-compatible.";
                validationErrorKind = ProviderValidationErrorKind.MissingApiKey;
                return false;
            }

            if (string.IsNullOrWhiteSpace(settings.LocalAiModelName))
            {
                validationError = "Thiếu model cho OpenAI-compatible.";
                validationErrorKind = ProviderValidationErrorKind.InvalidConfiguration;
                return false;
            }

            validationError = string.Empty;
            validationErrorKind = ProviderValidationErrorKind.None;
            return true;
        }

        var geminiKeys = GetDecryptedApiKeysFromSettings(settings);
        if (geminiKeys.Count == 0)
        {
            apiKeys = [];
            validationError = "Bạn chưa thêm Gemini API key.";
            validationErrorKind = ProviderValidationErrorKind.MissingApiKey;
            return false;
        }

        apiKeys = geminiKeys;
        validationError = string.Empty;
        validationErrorKind = ProviderValidationErrorKind.None;
        return true;
    }

    private async void ApiProviderRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        var provider = GetApiProviderFromUi();
        UpdateApiProviderUiState(provider);
        await AutoSaveNonApiSettingsAsync();
    }

    private void UpdateApiProviderUiState(string? provider)
    {
        var normalized = NormalizeApiProvider(provider);
        var useGemini = IsGeminiProvider(normalized);

        if (ApiProviderGeminiRadioButton.IsChecked != useGemini)
        {
            ApiProviderGeminiRadioButton.IsChecked = useGemini;
        }

        if (ApiProviderLocalAiRadioButton.IsChecked != !useGemini)
        {
            ApiProviderLocalAiRadioButton.IsChecked = !useGemini;
        }

        GeminiSectionBorder.Visibility = useGemini ? Visibility.Visible : Visibility.Collapsed;
        LocalAiSectionBorder.Visibility = useGemini ? Visibility.Collapsed : Visibility.Visible;
        UpdateLocalAiModeUiState(LocalAiUseCustomBaseUrlToggleSwitch.IsChecked == true);
    }

    private void UpdateLocalAiModeUiState(bool useCustomBaseUrl)
    {
        LocalAiBaseUrlPanel.Visibility = useCustomBaseUrl ? Visibility.Visible : Visibility.Collapsed;
        LocalAiApiKeyLabelTextBlock.Text = useCustomBaseUrl ? "API key (tùy chọn)" : "API key (bắt buộc)";
    }

    private async void LocalAiUseCustomBaseUrlToggleSwitch_OnChanged(object sender, RoutedEventArgs e)
    {
        UpdateLocalAiModeUiState(LocalAiUseCustomBaseUrlToggleSwitch.IsChecked == true);
        await AutoSaveNonApiSettingsAsync();
    }

    private async void SettingsControl_OnChanged(object sender, RoutedEventArgs e)
    {
        await AutoSaveNonApiSettingsAsync();
    }

    private async Task AutoSaveNonApiSettingsAsync()
    {
        if (!IsLoaded || _isApplyingSettings || _isAutoSavingSettings)
        {
            return;
        }

        _isAutoSavingSettings = true;
        try
        {
            var previousSettings = _state.Settings;
            var newSettings = BuildSettingsFromUi(includeApiKey: false);

            if (!TryRegisterHotkey(newSettings, out var error))
            {
                _isHotkeyRegistered = TryRegisterHotkey(previousSettings, out _);
                ApplySettingsToUi(previousSettings, includeApiKey: false);
                ShowStatus("Không áp dụng được hotkey", error, InfoBarSeverity.Error);
                return;
            }

            _isHotkeyRegistered = true;

            try
            {
                _startupRegistrationService.SetEnabled(newSettings.LaunchOnStartup);
            }
            catch (Exception ex)
            {
                _isHotkeyRegistered = TryRegisterHotkey(previousSettings, out _);
                ApplySettingsToUi(previousSettings, includeApiKey: false);
                ShowStatus("Không áp dụng được startup", ex.Message, InfoBarSeverity.Error);
                return;
            }

            try
            {
                ApplyTheme(newSettings.AppTheme);
            }
            catch (Exception ex)
            {
                _isHotkeyRegistered = TryRegisterHotkey(previousSettings, out _);
                ApplyTheme(previousSettings.AppTheme);
                ApplySettingsToUi(previousSettings, includeApiKey: false);
                ShowStatus("Không áp dụng được giao diện", ex.Message, InfoBarSeverity.Error);
                return;
            }

            _state.Settings = newSettings;
            await PersistStateAsync();
            ShowHotkeyPreview(newSettings);
            if (!HasActiveTranslation())
            {
                SetMainTranslationIdleStatus();
            }

            if (!HasActiveManualTranslation())
            {
                SetManualTranslationIdleStatus();
            }
        }
        finally
        {
            _isAutoSavingSettings = false;
        }
    }

    private async void SaveApiKeyButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var encryptedKeys = BuildEncryptedApiKeysFromUi();
            _state.Settings.GeminiApiKeysEncrypted = encryptedKeys;
            _state.Settings.GeminiApiKeyEncrypted = encryptedKeys.FirstOrDefault();
            await PersistStateAsync();
            if (encryptedKeys.Count == 0)
            {
                ShowStatus("Đã lưu API keys", "Chưa có API key nào hợp lệ để lưu.", InfoBarSeverity.Warning);
                return;
            }

            ShowStatus("Đã lưu API keys", $"Đã lưu {encryptedKeys.Count} Gemini API key.", InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            ShowStatus("Không lưu được API key", ex.Message, InfoBarSeverity.Error);
        }
    }

    private async void SaveLocalAiButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var localBaseUrl = NormalizeLocalAiBaseUrl(LocalAiBaseUrlTextBox.Text);
            var useCustomBaseUrl = LocalAiUseCustomBaseUrlToggleSwitch.IsChecked == true;
            var apiKey = LocalAiApiKeyTextBox.Text?.Trim();

            if (useCustomBaseUrl && !Uri.TryCreate(localBaseUrl, UriKind.Absolute, out _))
            {
                ShowStatus(
                    "Không lưu được OpenAI-compatible",
                    "Base URL tuỳ chỉnh không hợp lệ.",
                    InfoBarSeverity.Error
                );
                return;
            }

            if (!useCustomBaseUrl && string.IsNullOrWhiteSpace(apiKey))
            {
                ShowStatus(
                    "Không lưu được OpenAI-compatible",
                    "Khi tắt base URL tuỳ chỉnh, bạn cần nhập API key.",
                    InfoBarSeverity.Error
                );
                return;
            }

            _state.Settings.LocalAiBaseUrl = localBaseUrl;
            _state.Settings.LocalAiUseCustomBaseUrl = useCustomBaseUrl;
            _state.Settings.LocalAiModelName = NormalizeLocalAiModelName(LocalAiModelNameTextBox.Text);
            _state.Settings.LocalAiApiKeyEncrypted = BuildEncryptedLocalAiApiKeyFromPlainText(apiKey);
            _state.Settings.ActiveApiProvider = NormalizeApiProvider(GetApiProviderFromUi());
            await PersistStateAsync();
            ApplySettingsToUi(_state.Settings);
            ShowStatus("Đã lưu OpenAI-compatible", "Đã lưu cấu hình OpenAI-compatible.", InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            ShowStatus("Không lưu được OpenAI-compatible", ex.Message, InfoBarSeverity.Error);
        }
    }

    private async void TestLocalAiButton_Click(object sender, RoutedEventArgs e)
    {
        if (!await _translateLock.WaitAsync(0))
        {
            ShowStatus("Đang bận", "Vui lòng chờ tác vụ dịch hiện tại hoàn tất.", InfoBarSeverity.Warning);
            return;
        }

        try
        {
            var settings = BuildSettingsFromUi(includeApiKey: false);
            var localBaseUrl = ResolveLocalAiBaseUrl(settings);
            var localModel = NormalizeLocalAiModelName(settings.LocalAiModelName);
            var localApiKey = LocalAiApiKeyTextBox.Text?.Trim();

            if (!settings.LocalAiUseCustomBaseUrl && string.IsNullOrWhiteSpace(localApiKey))
            {
                ShowStatus(
                    "Thiếu API key",
                    "Khi tắt base URL tuỳ chỉnh, bạn cần nhập API key để test.",
                    InfoBarSeverity.Error
                );
                return;
            }

            if (settings.LocalAiUseCustomBaseUrl && !Uri.TryCreate(localBaseUrl, UriKind.Absolute, out _))
            {
                ShowStatus(
                    "Test OpenAI-compatible thất bại",
                    "Base URL tuỳ chỉnh không hợp lệ.",
                    InfoBarSeverity.Error
                );
                return;
            }

            _ = await _localAiTranslationService.TranslateAsync(
                "Xin chao",
                localBaseUrl,
                localApiKey,
                localModel,
                settings.TargetLanguage
            );

            ShowStatus(
                "Kết nối thành công",
                "OpenAI-compatible phản hồi thành công với cấu hình hiện tại.",
                InfoBarSeverity.Success
            );
        }
        catch (Exception ex)
        {
            ShowStatus("Test OpenAI-compatible thất bại", ex.Message, InfoBarSeverity.Error);
        }
        finally
        {
            _translateLock.Release();
        }
    }

    private void AddApiKeyRowButton_Click(object sender, RoutedEventArgs e)
    {
        ApiKeyInputItems.Add(new ApiKeyInputItem());
    }

    private void RemoveApiKeyRowButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: ApiKeyInputItem item })
        {
            return;
        }

        ApiKeyInputItems.Remove(item);
        if (ApiKeyInputItems.Count == 0)
        {
            ApiKeyInputItems.Add(new ApiKeyInputItem());
        }
    }

    private async void TestApiKeyRowButton_Click(object sender, RoutedEventArgs e)
    {
        if (!await _translateLock.WaitAsync(0))
        {
            ShowStatus("Đang bận", "Vui lòng chờ tác vụ dịch hiện tại hoàn tất.", InfoBarSeverity.Warning);
            return;
        }

        try
        {
            if (sender is not FrameworkElement { DataContext: ApiKeyInputItem item })
            {
                ShowStatus("Không test được key", "Không xác định được dòng API key.", InfoBarSeverity.Error);
                return;
            }

            var apiKey = item.PlainText?.Trim();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                ShowStatus("Thiếu API key", "Nhập API key vào dòng cần test.", InfoBarSeverity.Error);
                return;
            }

            var settings = BuildSettingsFromUi(includeApiKey: false);
            _ = await _translationService.TranslateAsync(
                "Xin chao",
                [apiKey],
                settings.GeminiModelName,
                settings.TargetLanguage
            );

            var index = ApiKeyInputItems.IndexOf(item) + 1;
            ShowStatus("Kết nối thành công", $"API key dòng {index} hợp lệ với model hiện tại.", InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            ShowStatus("Test thất bại", ex.Message, InfoBarSeverity.Error);
        }
        finally
        {
            _translateLock.Release();
        }
    }

    private async void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
    {
        HistoryRecords.Clear();
        _state.History.Clear();
        await PersistStateAsync();
        ShowStatus("Đã xóa lịch sử", "Toàn bộ lịch sử dịch đã được xóa.", InfoBarSeverity.Success);
    }

    private void CopySelectedHistoryButton_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListView.SelectedItem is not TranslationRecord record)
        {
            ShowStatus("Chưa chọn mục", "Chọn một bản ghi trong lịch sử trước khi copy.", InfoBarSeverity.Warning);
            return;
        }

        System.Windows.Clipboard.SetText(record.TranslatedText);
        ShowStatus("Đã copy", "Bản dịch đã chọn đã được copy vào clipboard.", InfoBarSeverity.Success);
    }

    private AppSettings BuildSettingsFromUi(bool includeApiKey = true)
    {
        var activeProvider = GetApiProviderFromUi();
        var model = NormalizeModelName(ModelNameTextBox.Text);
        var localAiBaseUrl = NormalizeLocalAiBaseUrl(LocalAiBaseUrlTextBox.Text);
        var localAiUseCustomBaseUrl = LocalAiUseCustomBaseUrlToggleSwitch.IsChecked == true;
        var localAiModelName = NormalizeLocalAiModelName(LocalAiModelNameTextBox.Text);
        var hotkeyKey = NormalizeHotkeyKey(HotkeyKeyComboBox.SelectedItem?.ToString());
        var targetLanguage = NormalizeTargetLanguage(TargetLanguageComboBox.SelectedValue?.ToString());
        var encryptedKeys = includeApiKey
            ? BuildEncryptedApiKeysFromUi()
            : GetEncryptedApiKeysFromSettings(_state.Settings);
        var localAiApiKeyEncrypted = includeApiKey
            ? BuildEncryptedLocalAiApiKeyFromUi()
            : _state.Settings.LocalAiApiKeyEncrypted;

        return new AppSettings
        {
            AppTheme = NormalizeAppTheme(AppThemeComboBox.SelectedValue?.ToString()),
            ActiveApiProvider = activeProvider,
            GeminiApiKeysEncrypted = encryptedKeys,
            GeminiApiKeyEncrypted = encryptedKeys.FirstOrDefault(),
            GeminiModelName = model,
            LocalAiBaseUrl = localAiBaseUrl,
            LocalAiUseCustomBaseUrl = localAiUseCustomBaseUrl,
            LocalAiApiKeyEncrypted = localAiApiKeyEncrypted,
            LocalAiModelName = localAiModelName,
            TargetLanguage = targetLanguage,
            CopyTranslationToClipboard = CopyToClipboardCheckBox.IsChecked == true,
            KeepRunningInBackgroundOnClose = KeepRunningInBackgroundOnCloseCheckBox.IsChecked == true,
            LaunchOnStartup = StartupWithWindowsCheckBox.IsChecked == true,
            HotkeyCtrl = HotkeyCtrlCheckBox.IsChecked == true,
            HotkeyShift = HotkeyShiftCheckBox.IsChecked == true,
            HotkeyAlt = HotkeyAltCheckBox.IsChecked == true,
            HotkeyWin = HotkeyWinCheckBox.IsChecked == true,
            HotkeyKey = hotkeyKey
        };
    }

    private async Task<string> TranslateWithSettingsAsync(
        string sourceText,
        AppSettings settings,
        IReadOnlyList<string> apiKeys,
        CancellationToken cancellationToken = default
    )
    {
        var provider = NormalizeApiProvider(settings.ActiveApiProvider);
        var modelName = GetActiveModelName(settings);

        ShowStatus(
            "Đang dịch",
            $"Provider: {GetProviderDisplayName(provider)} | Model: {modelName} | Ngôn ngữ đích: {GetTargetLanguageDisplayName(settings.TargetLanguage)}",
            InfoBarSeverity.Informational
        );

        if (IsLocalAiProvider(provider))
        {
            var localApiKey = GetDecryptedLocalAiApiKeyFromSettings(settings);
            return await _localAiTranslationService.TranslateAsync(
                sourceText,
                ResolveLocalAiBaseUrl(settings),
                localApiKey,
                settings.LocalAiModelName,
                settings.TargetLanguage,
                cancellationToken
            );
        }

        return await _translationService.TranslateAsync(
            sourceText,
            apiKeys,
            settings.GeminiModelName,
            settings.TargetLanguage,
            cancellationToken
        );
    }

    private bool TryRegisterHotkey(AppSettings settings, out string error)
    {
        error = string.Empty;

        if (!TryBuildHotkey(settings, out var modifiers, out var key, out error))
        {
            return false;
        }

        _hotkeyService?.Dispose();
        _hotkeyService = new GlobalHotkeyService(this, HotkeyId, modifiers, key);
        _hotkeyService.HotkeyPressed += HotkeyServiceOnHotkeyPressed;

        if (_hotkeyService.Register())
        {
            return true;
        }

        _hotkeyService.Dispose();
        _hotkeyService = null;
        error = $"Phím tắt {BuildHotkeyDisplay(settings)} có thể đang bị ứng dụng khác sử dụng.";
        return false;
    }

    private static bool TryBuildHotkey(AppSettings settings, out ModifierKeys modifiers, out Key key, out string error)
    {
        modifiers = ModifierKeys.None;
        key = Key.None;
        error = string.Empty;

        if (settings.HotkeyCtrl)
        {
            modifiers |= ModifierKeys.Control;
        }

        if (settings.HotkeyShift)
        {
            modifiers |= ModifierKeys.Shift;
        }

        if (settings.HotkeyAlt)
        {
            modifiers |= ModifierKeys.Alt;
        }

        if (settings.HotkeyWin)
        {
            modifiers |= ModifierKeys.Windows;
        }

        if (modifiers == ModifierKeys.None)
        {
            error = "Hotkey cần ít nhất một phím bổ trợ (Ctrl/Shift/Alt/Win).";
            return false;
        }

        if (!Enum.TryParse(settings.HotkeyKey, true, out key) || key == Key.None)
        {
            error = "Phím chính của hotkey không hợp lệ.";
            return false;
        }

        return true;
    }

    private void AddToHistory(string sourceText, string translatedText, string modelName)
    {
        var record = new TranslationRecord
        {
            Timestamp = DateTimeOffset.Now,
            SourceText = sourceText,
            TranslatedText = translatedText,
            ModelName = modelName
        };

        HistoryRecords.Insert(0, record);
        _state.History.Insert(0, record);

        const int maxRecords = 500;
        while (HistoryRecords.Count > maxRecords)
        {
            HistoryRecords.RemoveAt(HistoryRecords.Count - 1);
        }

        while (_state.History.Count > maxRecords)
        {
            _state.History.RemoveAt(_state.History.Count - 1);
        }
    }

    private async Task PersistStateAsync()
    {
        await _stateSaveLock.WaitAsync();
        try
        {
            await _stateStore.SaveAsync(_state);
        }
        finally
        {
            _stateSaveLock.Release();
        }
    }

    private void ShowProviderValidationError(string validationError, ProviderValidationErrorKind validationErrorKind)
    {
        if (validationErrorKind == ProviderValidationErrorKind.MissingApiKey)
        {
            ShowMissingApiKeySnackbar(validationError);
            return;
        }

        ShowStatus("Thiếu cấu hình API", validationError, InfoBarSeverity.Error);
    }

    private void ShowMissingApiKeySnackbar(string validationError)
    {
        var message = string.IsNullOrWhiteSpace(validationError)
            ? "Bạn chưa thiết lập API key."
            : validationError;

        try
        {
            var presenter = _snackbarService.GetSnackbarPresenter();
            if (presenter is null)
            {
                throw new InvalidOperationException("Snackbar presenter is not available.");
            }

            var openSettingsButton = new Wpf.Ui.Controls.Button
            {
                Content = "Mở Cài đặt",
                Appearance = ControlAppearance.Secondary,
                Icon = new SymbolIcon { Symbol = SymbolRegular.Settings24 },
                Margin = new Thickness(10, 0, 0, 0),
                Padding = new Thickness(10, 2, 10, 2),
                MinHeight = 28
            };
            openSettingsButton.Click += (_, _) => SwitchToTab(3, SettingsNavItem);

            var content = new StackPanel { Orientation = Orientation.Horizontal };
            content.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = message,
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            });
            content.Children.Add(openSettingsButton);

            var snackbar = new Snackbar(presenter)
            {
                Title = "Chưa thiết lập API key",
                Content = content,
                Appearance = ControlAppearance.Danger,
                Icon = new SymbolIcon { Symbol = SymbolRegular.ErrorCircle24 },
                Timeout = TimeSpan.FromSeconds(7),
                IsCloseButtonEnabled = true
            };

            _ = presenter.ImmediatelyDisplay(snackbar);
        }
        catch
        {
            ShowStatus(
                "Chưa thiết lập API key",
                $"{message} Vào tab Cài đặt để thêm API key.",
                InfoBarSeverity.Error
            );
        }
    }

    private void ShowStatus(string title, string message, InfoBarSeverity severity)
    {
        var (appearance, icon, timeout) = MapSnackbarStyle(severity);
        _snackbarService.Show(title, message, appearance, icon, timeout);
    }

    private static (ControlAppearance Appearance, IconElement Icon, TimeSpan Timeout) MapSnackbarStyle(InfoBarSeverity severity)
    {
        return severity switch
        {
            InfoBarSeverity.Success => (ControlAppearance.Success, new SymbolIcon { Symbol = SymbolRegular.Checkmark24 }, TimeSpan.FromSeconds(3)),
            InfoBarSeverity.Error => (ControlAppearance.Danger, new SymbolIcon { Symbol = SymbolRegular.ErrorCircle24 }, TimeSpan.FromSeconds(4)),
            InfoBarSeverity.Warning => (ControlAppearance.Caution, new SymbolIcon { Symbol = SymbolRegular.Clock24 }, TimeSpan.FromSeconds(3)),
            _ => (ControlAppearance.Info, new SymbolIcon { Symbol = SymbolRegular.Clock24 }, TimeSpan.FromSeconds(3))
        };
    }

    private void ApplySettingsToUi(AppSettings settings, bool includeApiKey = true)
    {
        _isApplyingSettings = true;
        try
        {
            var provider = NormalizeApiProvider(settings.ActiveApiProvider);
            if (includeApiKey)
            {
                ApiKeyInputItems.Clear();
                foreach (var key in GetDecryptedApiKeysFromSettings(settings))
                {
                    ApiKeyInputItems.Add(new ApiKeyInputItem { PlainText = key });
                }

                if (ApiKeyInputItems.Count == 0)
                {
                    ApiKeyInputItems.Add(new ApiKeyInputItem());
                }

                LocalAiApiKeyTextBox.Text = GetDecryptedLocalAiApiKeyFromSettings(settings) ?? string.Empty;
            }

            ApiProviderGeminiRadioButton.IsChecked = IsGeminiProvider(provider);
            ApiProviderLocalAiRadioButton.IsChecked = IsLocalAiProvider(provider);
            ModelNameTextBox.Text = settings.GeminiModelName;
            LocalAiBaseUrlTextBox.Text = NormalizeLocalAiBaseUrl(settings.LocalAiBaseUrl);
            LocalAiUseCustomBaseUrlToggleSwitch.IsChecked = settings.LocalAiUseCustomBaseUrl;
            LocalAiModelNameTextBox.Text = NormalizeLocalAiModelName(settings.LocalAiModelName);
            TargetLanguageComboBox.SelectedValue = NormalizeTargetLanguage(settings.TargetLanguage);
            if (!_isManualTargetLanguageInitialized || ManualTargetLanguageComboBox.SelectedValue is null)
            {
                ManualTargetLanguageComboBox.SelectedValue = NormalizeTargetLanguage(settings.TargetLanguage);
                _isManualTargetLanguageInitialized = true;
            }
            AppThemeComboBox.SelectedValue = NormalizeAppTheme(settings.AppTheme);
            CopyToClipboardCheckBox.IsChecked = settings.CopyTranslationToClipboard;
            KeepRunningInBackgroundOnCloseCheckBox.IsChecked = settings.KeepRunningInBackgroundOnClose;
            StartupWithWindowsCheckBox.IsChecked = settings.LaunchOnStartup;

            HotkeyCtrlCheckBox.IsChecked = settings.HotkeyCtrl;
            HotkeyShiftCheckBox.IsChecked = settings.HotkeyShift;
            HotkeyAltCheckBox.IsChecked = settings.HotkeyAlt;
            HotkeyWinCheckBox.IsChecked = settings.HotkeyWin;
            HotkeyKeyComboBox.SelectedItem = settings.HotkeyKey;

            ShowHotkeyPreview(settings);
            UpdateApiProviderUiState(provider);
        }
        finally
        {
            _isApplyingSettings = false;
        }
    }

    private void ShowHotkeyPreview(AppSettings settings)
    {
        HotkeyPreviewTextBlock.Text = $"Hotkey hiện tại: {BuildHotkeyDisplay(settings)}";
    }

    private static string BuildHotkeyDisplay(AppSettings settings)
    {
        var parts = new List<string>(5);
        if (settings.HotkeyCtrl)
        {
            parts.Add("Ctrl");
        }

        if (settings.HotkeyShift)
        {
            parts.Add("Shift");
        }

        if (settings.HotkeyAlt)
        {
            parts.Add("Alt");
        }

        if (settings.HotkeyWin)
        {
            parts.Add("Win");
        }

        parts.Add(settings.HotkeyKey);
        return string.Join("+", parts);
    }

    private static string NormalizeApiProvider(string? provider)
    {
        if (string.IsNullOrWhiteSpace(provider))
        {
            return ApiProviderGemini;
        }

        var normalized = provider.Trim();
        if (string.Equals(normalized, ApiProviderLocalAi, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "Local AI", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "LocalAI", StringComparison.OrdinalIgnoreCase))
        {
            return ApiProviderLocalAi;
        }

        return ApiProviderGemini;
    }

    private static bool IsLocalAiProvider(string? provider)
    {
        return string.Equals(NormalizeApiProvider(provider), ApiProviderLocalAi, StringComparison.Ordinal);
    }

    private static bool IsGeminiProvider(string? provider)
    {
        return string.Equals(NormalizeApiProvider(provider), ApiProviderGemini, StringComparison.Ordinal);
    }

    private static string NormalizeLocalAiBaseUrl(string? baseUrl)
    {
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            return LegacyDefaultLocalAiBaseUrl;
        }

        return baseUrl.Trim();
    }

    private static string ResolveLocalAiBaseUrl(AppSettings settings)
    {
        return settings.LocalAiUseCustomBaseUrl
            ? NormalizeLocalAiBaseUrl(settings.LocalAiBaseUrl)
            : DefaultLocalAiBaseUrl;
    }

    private static string NormalizeLocalAiModelName(string? modelName)
    {
        if (string.IsNullOrWhiteSpace(modelName))
        {
            return DefaultLocalAiModel;
        }

        return modelName.Trim();
    }

    private static string GetProviderDisplayName(string? provider)
    {
        return IsLocalAiProvider(provider) ? "OpenAI-compatible" : "Gemini";
    }

    private static string GetActiveModelName(AppSettings settings)
    {
        return IsLocalAiProvider(settings.ActiveApiProvider)
            ? NormalizeLocalAiModelName(settings.LocalAiModelName)
            : NormalizeModelName(settings.GeminiModelName);
    }

    private static string NormalizeModelName(string? model)
    {
        if (string.IsNullOrWhiteSpace(model))
        {
            return DefaultModels[0];
        }

        var trimmed = model.Trim();
        if (string.Equals(trimmed, "gemini-2.0-flash-lite", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(trimmed, "gemini-2.0-flash-lite-001", StringComparison.OrdinalIgnoreCase))
        {
            return "gemini-flash-lite-latest";
        }

        return trimmed;
    }

    private static string NormalizeHotkeyKey(string? key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return "E";
        }

        var normalized = key.Trim().ToUpperInvariant();
        return HotkeyKeys.Contains(normalized) ? normalized : "E";
    }

    private static string NormalizeTargetLanguage(string? targetLanguage)
    {
        if (string.IsNullOrWhiteSpace(targetLanguage))
        {
            return DefaultTargetLanguage;
        }

        var normalized = targetLanguage.Trim();
        var matched = SupportedLanguages.FirstOrDefault(x =>
            string.Equals(x.PromptName, normalized, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(x.DisplayName, normalized, StringComparison.OrdinalIgnoreCase));

        return matched?.PromptName ?? DefaultTargetLanguage;
    }

    private static string NormalizeAppTheme(string? appTheme)
    {
        return string.Equals(appTheme?.Trim(), "Dark", StringComparison.OrdinalIgnoreCase)
            ? "Dark"
            : "Light";
    }

    private static bool IsDarkTheme(string? appTheme)
    {
        return string.Equals(NormalizeAppTheme(appTheme), "Dark", StringComparison.Ordinal);
    }

    private static ApplicationTheme ToApplicationTheme(string? appTheme)
    {
        return IsDarkTheme(appTheme) ? ApplicationTheme.Dark : ApplicationTheme.Light;
    }

    private void ApplyTheme(string? appTheme)
    {
        ApplicationThemeManager.Apply(ToApplicationTheme(appTheme));
        SetResourceReference(BackgroundProperty, "ApplicationBackgroundBrush");
        MainRootGrid.SetResourceReference(Panel.BackgroundProperty, "ApplicationBackgroundBrush");
        MainContentGrid.SetResourceReference(Panel.BackgroundProperty, "ApplicationBackgroundBrush");
        ContentTabControl.ClearValue(Control.BackgroundProperty);
    }

    private static string GetTargetLanguageDisplayName(string? targetLanguage)
    {
        var normalized = NormalizeTargetLanguage(targetLanguage);
        return SupportedLanguages.FirstOrDefault(x => string.Equals(x.PromptName, normalized, StringComparison.OrdinalIgnoreCase))
                   ?.DisplayName
               ?? DefaultTargetLanguage;
    }

    private bool EnsureSettingsCompatibility(AppSettings settings)
    {
        var changed = false;
        var normalizedProvider = NormalizeApiProvider(settings.ActiveApiProvider);
        if (!string.Equals(settings.ActiveApiProvider, normalizedProvider, StringComparison.Ordinal))
        {
            settings.ActiveApiProvider = normalizedProvider;
            changed = true;
        }

        var normalizedApiKeys = GetEncryptedApiKeysFromSettings(settings);
        if (settings.GeminiApiKeysEncrypted is null ||
            !settings.GeminiApiKeysEncrypted.SequenceEqual(normalizedApiKeys, StringComparer.Ordinal))
        {
            settings.GeminiApiKeysEncrypted = normalizedApiKeys;
            changed = true;
        }

        var firstApiKey = normalizedApiKeys.FirstOrDefault();
        if (!string.Equals(settings.GeminiApiKeyEncrypted, firstApiKey, StringComparison.Ordinal))
        {
            settings.GeminiApiKeyEncrypted = firstApiKey;
            changed = true;
        }

        var normalizedModel = NormalizeModelName(settings.GeminiModelName);
        if (!string.Equals(normalizedModel, settings.GeminiModelName, StringComparison.Ordinal))
        {
            settings.GeminiModelName = normalizedModel;
            changed = true;
        }

        var useCustomBaseUrl = settings.LocalAiUseCustomBaseUrl;
        var existingLocalAiBaseUrl = settings.LocalAiBaseUrl?.Trim();
        if (!useCustomBaseUrl &&
            !string.IsNullOrWhiteSpace(existingLocalAiBaseUrl) &&
            !string.Equals(existingLocalAiBaseUrl, DefaultLocalAiBaseUrl, StringComparison.OrdinalIgnoreCase))
        {
            // Migrate older settings (previously always custom URL) without breaking existing behavior.
            useCustomBaseUrl = true;
        }

        if (settings.LocalAiUseCustomBaseUrl != useCustomBaseUrl)
        {
            settings.LocalAiUseCustomBaseUrl = useCustomBaseUrl;
            changed = true;
        }

        var normalizedLocalAiBaseUrl = NormalizeLocalAiBaseUrl(settings.LocalAiBaseUrl);
        if (!string.Equals(normalizedLocalAiBaseUrl, settings.LocalAiBaseUrl, StringComparison.Ordinal))
        {
            settings.LocalAiBaseUrl = normalizedLocalAiBaseUrl;
            changed = true;
        }

        var normalizedLocalAiModelName = NormalizeLocalAiModelName(settings.LocalAiModelName);
        if (!string.Equals(normalizedLocalAiModelName, settings.LocalAiModelName, StringComparison.Ordinal))
        {
            settings.LocalAiModelName = normalizedLocalAiModelName;
            changed = true;
        }

        var normalizedTargetLanguage = NormalizeTargetLanguage(settings.TargetLanguage);
        if (!string.Equals(normalizedTargetLanguage, settings.TargetLanguage, StringComparison.Ordinal))
        {
            settings.TargetLanguage = normalizedTargetLanguage;
            changed = true;
        }

        var normalizedAppTheme = NormalizeAppTheme(settings.AppTheme);
        if (!string.Equals(normalizedAppTheme, settings.AppTheme, StringComparison.Ordinal))
        {
            settings.AppTheme = normalizedAppTheme;
            changed = true;
        }

        var normalizedKey = NormalizeHotkeyKey(settings.HotkeyKey);
        if (!string.Equals(normalizedKey, settings.HotkeyKey, StringComparison.Ordinal))
        {
            settings.HotkeyKey = normalizedKey;
            changed = true;
        }

        if (!settings.HotkeyCtrl && !settings.HotkeyShift && !settings.HotkeyAlt && !settings.HotkeyWin)
        {
            settings.HotkeyCtrl = true;
            settings.HotkeyShift = true;
            changed = true;
        }

        return changed;
    }

    private void InitializeLanguageOptions()
    {
        if (TargetLanguageOptions.Count > 0)
        {
            return;
        }

        foreach (var language in SupportedLanguages)
        {
            TargetLanguageOptions.Add(language);
        }
    }

    private void InitializeThemeOptions()
    {
        if (ThemeOptions.Count > 0)
        {
            return;
        }

        ThemeOptions.Add(new ThemeOptionItem("Sáng", "Light"));
        ThemeOptions.Add(new ThemeOptionItem("Tối", "Dark"));
    }

    private void InitializeHotkeyOptions()
    {
        if (HotkeyKeyOptions.Count > 0)
        {
            return;
        }

        foreach (var key in HotkeyKeys)
        {
            HotkeyKeyOptions.Add(key);
        }
    }

    private static string[] BuildHotkeyKeys()
    {
        var keys = new List<string>(26 + 12);

        for (var c = 'A'; c <= 'Z'; c++)
        {
            keys.Add(c.ToString());
        }

        for (var i = 1; i <= 12; i++)
        {
            keys.Add($"F{i}");
        }

        return keys.ToArray();
    }

    private void GeminiApiKeyLink_OnRequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true
            });
            e.Handled = true;
        }
        catch (Exception ex)
        {
            ShowStatus("Không mở được liên kết", ex.Message, InfoBarSeverity.Warning);
        }
    }

}
