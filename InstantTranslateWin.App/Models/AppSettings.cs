namespace InstantTranslateWin.App.Models;

public sealed class AppSettings
{
    public string AppTheme { get; set; } = "Light";

    public string ActiveApiProvider { get; set; } = "Gemini";

    public List<string> GeminiApiKeysEncrypted { get; set; } = [];

    public string? GeminiApiKeyEncrypted { get; set; }

    public string GeminiModelName { get; set; } = "gemini-flash-lite-latest";

    public string LocalAiBaseUrl { get; set; } = "https://api.openai.com/v1";

    public bool LocalAiUseCustomBaseUrl { get; set; }

    public string? LocalAiApiKeyEncrypted { get; set; }

    public string LocalAiModelName { get; set; } = "gpt-4o-mini";

    public string TargetLanguage { get; set; } = "English";

    public bool CopyTranslationToClipboard { get; set; } = true;

    public bool KeepRunningInBackgroundOnClose { get; set; } = true;

    public bool LaunchOnStartup { get; set; }

    public bool HotkeyCtrl { get; set; } = true;

    public bool HotkeyShift { get; set; } = true;

    public bool HotkeyAlt { get; set; }

    public bool HotkeyWin { get; set; }

    public string HotkeyKey { get; set; } = "E";

    public bool QuickInputHotkeyCtrl { get; set; } = true;

    public bool QuickInputHotkeyShift { get; set; } = true;

    public bool QuickInputHotkeyAlt { get; set; }

    public bool QuickInputHotkeyWin { get; set; }

    public string QuickInputHotkeyKey { get; set; } = "H";

    public string QuickInputInputLanguage { get; set; } = QuickInputTypingOptions.InputLanguageVietnamese;

    public string QuickInputVietnameseTypingStyle { get; set; } = QuickInputTypingOptions.VietnameseTypingStyleTelex;
}
