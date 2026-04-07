using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using InstantTranslateWin.App.Models;
using Wpf.Ui.Controls;

namespace InstantTranslateWin.App;

public partial class QuickInputPopup : FluentWindow
{
    private const uint MonitorDefaultToNearest = 0x00000002;
    private const double EdgeMargin = 12;
    private const double CursorGap = 32;

    private bool _allowClose;
    private bool _isApplyingInputOptions;

    public event EventHandler<string>? SubmitRequested;
    public event EventHandler<QuickInputPopupSettingsChangedEventArgs>? InputOptionsChanged;

    public bool IsInputFocused => InputTextBox.IsKeyboardFocusWithin;
    public string CurrentText => InputTextBox.Text ?? string.Empty;

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

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromPoint(NativePoint pt, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref NativeMonitorInfo lpmi);

    public QuickInputPopup()
    {
        InitializeComponent();
        ApplyInputOptions(
            QuickInputTypingOptions.InputLanguageVietnamese,
            QuickInputTypingOptions.VietnameseTypingStyleTelex
        );
    }

    public void ShowNearCursor(bool clearText = true)
    {
        if (clearText)
        {
            InputTextBox.Text = string.Empty;
        }

        if (!IsVisible)
        {
            Show();
            UpdateLayout();
        }

        PositionNearCursor();
    }

    public void HidePopup()
    {
        if (IsVisible)
        {
            Hide();
        }
    }

    public void CloseForShutdown()
    {
        _allowClose = true;
        Close();
    }

    public void ApplyInputOptions(string inputLanguage, string vietnameseTypingStyle)
    {
        _isApplyingInputOptions = true;
        try
        {
            var normalizedInputLanguage = NormalizeInputLanguage(inputLanguage);
            var normalizedTypingStyle = NormalizeVietnameseTypingStyle(vietnameseTypingStyle);

            SetComboBoxSelectedByTag(InputLanguageComboBox, normalizedInputLanguage);
            SetComboBoxSelectedByTag(VietnameseTypingStyleComboBox, normalizedTypingStyle);
            UpdateVietnameseTypingPanelVisibility(normalizedInputLanguage);
        }
        finally
        {
            _isApplyingInputOptions = false;
        }
    }

    public void InsertMirroredText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var currentText = InputTextBox.Text ?? string.Empty;
        var (selectionStart, selectionLength) = GetSafeSelection(currentText);
        var nextText = selectionLength > 0
            ? currentText.Remove(selectionStart, selectionLength).Insert(selectionStart, text)
            : currentText.Insert(selectionStart, text);
        SetTextAndCaret(nextText, selectionStart + text.Length);
    }

    public void ApplyMirroredBackspace()
    {
        var currentText = InputTextBox.Text ?? string.Empty;
        var (selectionStart, selectionLength) = GetSafeSelection(currentText);
        if (selectionLength > 0)
        {
            var nextText = currentText.Remove(selectionStart, selectionLength);
            SetTextAndCaret(nextText, selectionStart);
            return;
        }

        var caret = Math.Clamp(InputTextBox.CaretIndex, 0, currentText.Length);
        if (caret <= 0)
        {
            return;
        }

        var next = currentText.Remove(caret - 1, 1);
        SetTextAndCaret(next, caret - 1);
    }

    public void ApplyMirroredDelete()
    {
        var currentText = InputTextBox.Text ?? string.Empty;
        var (selectionStart, selectionLength) = GetSafeSelection(currentText);
        if (selectionLength > 0)
        {
            var nextText = currentText.Remove(selectionStart, selectionLength);
            SetTextAndCaret(nextText, selectionStart);
            return;
        }

        var caret = Math.Clamp(InputTextBox.CaretIndex, 0, currentText.Length);
        if (caret >= currentText.Length || currentText.Length == 0)
        {
            return;
        }

        var next = currentText.Remove(caret, 1);
        SetTextAndCaret(next, caret);
    }

    public void MoveCaretLeft()
    {
        var currentText = InputTextBox.Text ?? string.Empty;
        var (selectionStart, selectionLength) = GetSafeSelection(currentText);
        var caret = Math.Clamp(InputTextBox.CaretIndex, 0, currentText.Length);
        var targetCaret = selectionLength > 0 ? selectionStart : Math.Max(0, caret - 1);
        InputTextBox.Select(targetCaret, 0);
    }

    public void MoveCaretRight()
    {
        var currentText = InputTextBox.Text ?? string.Empty;
        var (selectionStart, selectionLength) = GetSafeSelection(currentText);
        var caret = Math.Clamp(InputTextBox.CaretIndex, 0, currentText.Length);
        var targetCaret = selectionLength > 0
            ? selectionStart + selectionLength
            : Math.Min(currentText.Length, caret + 1);
        InputTextBox.Select(targetCaret, 0);
    }

    public void MoveCaretHome()
    {
        InputTextBox.Select(0, 0);
    }

    public void MoveCaretEnd()
    {
        var end = (InputTextBox.Text ?? string.Empty).Length;
        InputTextBox.Select(end, 0);
    }

    public void SelectAllMirroredText()
    {
        InputTextBox.SelectAll();
    }

    public void SubmitFromMirroredInput()
    {
        SubmitInput();
    }

    public void SetMirroredText(string text)
    {
        var next = text ?? string.Empty;
        if (string.Equals(CurrentText, next, StringComparison.Ordinal))
        {
            return;
        }

        SetTextAndCaret(next, next.Length);
    }

    private void TranslateButton_OnClick(object sender, RoutedEventArgs e)
    {
        SubmitInput();
    }

    private void CloseButton_OnClick(object sender, RoutedEventArgs e)
    {
        HidePopup();
    }

    private void InputSettingsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var isVisible = InputSettingsPanel.Visibility == Visibility.Visible;
        InputSettingsPanel.Visibility = isVisible ? Visibility.Collapsed : Visibility.Visible;
        InputSettingsButton.Appearance = isVisible ? ControlAppearance.Secondary : ControlAppearance.Primary;
    }

    private void InputLanguageComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isApplyingInputOptions)
        {
            return;
        }

        var inputLanguage = GetSelectedComboTag(
            InputLanguageComboBox,
            QuickInputTypingOptions.InputLanguageVietnamese
        );
        UpdateVietnameseTypingPanelVisibility(inputLanguage);
        RaiseInputOptionsChanged();
    }

    private void VietnameseTypingStyleComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isApplyingInputOptions)
        {
            return;
        }

        RaiseInputOptionsChanged();
    }

    private void InputTextBox_OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            return;
        }

        if (e.Key != Key.Escape)
        {
            return;
        }

        e.Handled = true;
        HidePopup();
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape)
        {
            return;
        }

        e.Handled = true;
        HidePopup();
    }

    private void DragHandle_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed)
        {
            return;
        }

        if (e.OriginalSource is DependencyObject source)
        {
            if (FindVisualParent<ButtonBase>(source) is not null ||
                FindVisualParent<ComboBox>(source) is not null)
            {
                return;
            }
        }

        try
        {
            DragMove();
        }
        catch
        {
            // Ignore drag failures.
        }
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!_allowClose)
        {
            e.Cancel = true;
            Hide();
            return;
        }

        base.OnClosing(e);
    }

    private void SubmitInput()
    {
        var input = InputTextBox.Text;
        SubmitRequested?.Invoke(this, input);
    }

    private (int SelectionStart, int SelectionLength) GetSafeSelection(string currentText)
    {
        var safeStart = Math.Clamp(InputTextBox.SelectionStart, 0, currentText.Length);
        var safeLength = Math.Clamp(InputTextBox.SelectionLength, 0, currentText.Length - safeStart);
        return (safeStart, safeLength);
    }

    private void SetTextAndCaret(string text, int caretIndex)
    {
        InputTextBox.Text = text;
        var safeCaret = Math.Clamp(caretIndex, 0, text.Length);
        InputTextBox.Select(safeCaret, 0);
    }

    private void RaiseInputOptionsChanged()
    {
        var inputLanguage = NormalizeInputLanguage(
            GetSelectedComboTag(InputLanguageComboBox, QuickInputTypingOptions.InputLanguageVietnamese)
        );
        var typingStyle = NormalizeVietnameseTypingStyle(
            GetSelectedComboTag(VietnameseTypingStyleComboBox, QuickInputTypingOptions.VietnameseTypingStyleTelex)
        );
        InputOptionsChanged?.Invoke(this, new QuickInputPopupSettingsChangedEventArgs(inputLanguage, typingStyle));
    }

    private static string GetSelectedComboTag(ComboBox comboBox, string fallback)
    {
        if (comboBox.SelectedItem is ComboBoxItem item && item.Tag is string tag && !string.IsNullOrWhiteSpace(tag))
        {
            return tag;
        }

        return fallback;
    }

    private static void SetComboBoxSelectedByTag(ComboBox comboBox, string tag)
    {
        for (var index = 0; index < comboBox.Items.Count; index++)
        {
            if (comboBox.Items[index] is ComboBoxItem item &&
                item.Tag is string value &&
                string.Equals(value, tag, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedIndex = index;
                return;
            }
        }
    }

    private void UpdateVietnameseTypingPanelVisibility(string inputLanguage)
    {
        VietnameseTypingStylePanel.Visibility = string.Equals(
            NormalizeInputLanguage(inputLanguage),
            QuickInputTypingOptions.InputLanguageVietnamese,
            StringComparison.Ordinal
        )
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private static string NormalizeInputLanguage(string? inputLanguage)
    {
        return string.Equals(inputLanguage, QuickInputTypingOptions.InputLanguageOther, StringComparison.OrdinalIgnoreCase)
            ? QuickInputTypingOptions.InputLanguageOther
            : QuickInputTypingOptions.InputLanguageVietnamese;
    }

    private static string NormalizeVietnameseTypingStyle(string? typingStyle)
    {
        if (
            string.Equals(
                typingStyle,
                QuickInputTypingOptions.VietnameseTypingStyleViqr,
                StringComparison.OrdinalIgnoreCase
            )
        )
        {
            return QuickInputTypingOptions.VietnameseTypingStyleViqr;
        }

        return string.Equals(
            typingStyle,
            QuickInputTypingOptions.VietnameseTypingStyleVni,
            StringComparison.OrdinalIgnoreCase
        )
            ? QuickInputTypingOptions.VietnameseTypingStyleVni
            : QuickInputTypingOptions.VietnameseTypingStyleTelex;
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

    private void PositionNearCursor()
    {
        var hasCursor = GetCursorPos(out var cursorPoint);
        var cursorDevicePoint = hasCursor ? new Point(cursorPoint.X, cursorPoint.Y) : new Point(0, 0);
        var transformFromDevice = GetTransformFromDevice();
        var cursorPointDip = transformFromDevice.Transform(cursorDevicePoint);
        var workArea = hasCursor ? ResolveCurrentMonitorWorkArea(cursorPoint) : SystemParameters.WorkArea;

        var width = ResolveWindowWidth();
        var height = ResolveWindowHeight();

        var targetLeft = cursorPointDip.X - (width / 2);
        var targetTop = cursorPointDip.Y - height - CursorGap;
        if (targetTop < workArea.Top + EdgeMargin)
        {
            targetTop = cursorPointDip.Y + CursorGap;
        }

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

        Left = Math.Clamp(targetLeft, minLeft, maxLeft);
        Top = Math.Clamp(targetTop, minTop, maxTop);
    }

    private Matrix GetTransformFromDevice()
    {
        var source = PresentationSource.FromVisual(this);
        return source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
    }

    private Rect ResolveCurrentMonitorWorkArea(NativePoint cursorPoint)
    {
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

        var transformFromDevice = GetTransformFromDevice();
        var topLeft = transformFromDevice.Transform(new Point(monitorInfo.Work.Left, monitorInfo.Work.Top));
        var bottomRight = transformFromDevice.Transform(new Point(monitorInfo.Work.Right, monitorInfo.Work.Bottom));
        return new Rect(topLeft, bottomRight);
    }

    private double ResolveWindowWidth()
    {
        var width = ActualWidth > 0 ? ActualWidth : Width;
        if (double.IsNaN(width) || width <= 0)
        {
            width = 500;
        }

        return width;
    }

    private double ResolveWindowHeight()
    {
        var height = ActualHeight > 0 ? ActualHeight : Height;
        if (double.IsNaN(height) || height <= 0)
        {
            height = 74;
        }

        return height;
    }
}

public sealed class QuickInputPopupSettingsChangedEventArgs : EventArgs
{
    public string InputLanguage { get; }
    public string VietnameseTypingStyle { get; }

    public QuickInputPopupSettingsChangedEventArgs(string inputLanguage, string vietnameseTypingStyle)
    {
        InputLanguage = inputLanguage;
        VietnameseTypingStyle = vietnameseTypingStyle;
    }
}
