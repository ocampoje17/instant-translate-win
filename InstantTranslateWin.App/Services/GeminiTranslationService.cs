using Google.GenAI;
using Google.GenAI.Types;

namespace InstantTranslateWin.App.Services;

public sealed class GeminiTranslationService : IDisposable
{
    private readonly object _roundRobinLock = new();
    private int _lastApiKeyIndex = -1;

    public async Task<string> TranslateAsync(
        string sourceText,
        IReadOnlyList<string> apiKeys,
        string modelName,
        string targetLanguage,
        CancellationToken cancellationToken = default
    )
    {
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            throw new InvalidOperationException("Không có text để dịch.");
        }

        var normalizedApiKeys = (apiKeys ?? [])
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToList();

        if (normalizedApiKeys.Count == 0)
        {
            throw new InvalidOperationException("Thiếu Gemini API key.");
        }

        if (string.IsNullOrWhiteSpace(targetLanguage))
        {
            targetLanguage = "English";
        }

        var prompt =
            $"Translate the following text to natural {targetLanguage}. Return only the translation without additional explanation.\n\n"
            + sourceText;

        var config = new GenerateContentConfig
        {
            Temperature = 0.2f,
            ThinkingConfig = new ThinkingConfig
            {
                ThinkingBudget = 0
            }
        };

        Exception? lastError = null;
        const int maxAttempts = 2; // first attempt + one retry

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var apiKey = GetNextApiKey(normalizedApiKeys);
                using var client = new Client(apiKey: apiKey);
                var response = await client.Models.GenerateContentAsync(
                    modelName,
                    prompt,
                    config,
                    cancellationToken: cancellationToken
                );
                var translated = response.Text?.Trim();

                if (string.IsNullOrWhiteSpace(translated))
                {
                    throw new InvalidOperationException("Gemini API không trả về nội dung dịch.");
                }

                return translated;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex) when (attempt < maxAttempts)
            {
                lastError = ex;
                ErrorFileLogger.LogMessage(
                    "GeminiTranslationService.TranslateAsync.Retry",
                    $"Attempt {attempt}/{maxAttempts} failed: {ex.GetType().Name}: {ex.Message}"
                );
                await Task.Delay(TimeSpan.FromMilliseconds(450), cancellationToken);
            }
            catch (Exception ex)
            {
                lastError = ex;
                ErrorFileLogger.LogException("GeminiTranslationService.TranslateAsync.FinalAttempt", ex);
            }
        }

        throw lastError ?? new InvalidOperationException("Yêu cầu dịch thất bại.");
    }

    private string GetNextApiKey(IReadOnlyList<string> apiKeys)
    {
        lock (_roundRobinLock)
        {
            _lastApiKeyIndex = (_lastApiKeyIndex + 1) % apiKeys.Count;
            return apiKeys[_lastApiKeyIndex];
        }
    }

    public void Dispose()
    {
        // No shared resource to dispose because client is created per request.
    }
}
