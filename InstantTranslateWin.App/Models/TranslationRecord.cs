using System.Text.Json.Serialization;

namespace InstantTranslateWin.App.Models;

public sealed class TranslationRecord
{
    public DateTimeOffset Timestamp { get; set; } = DateTimeOffset.Now;

    public string SourceText { get; set; } = string.Empty;

    public string TranslatedText { get; set; } = string.Empty;

    public string ModelName { get; set; } = string.Empty;

    [JsonIgnore]
    public string TimestampDisplay => Timestamp.LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss");
}
