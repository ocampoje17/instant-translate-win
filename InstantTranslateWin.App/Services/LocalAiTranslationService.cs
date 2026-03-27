using OpenAI;
using OpenAI.Chat;
using System.ClientModel;

namespace InstantTranslateWin.App.Services;

public sealed class LocalAiTranslationService : IDisposable
{
    public async Task<string> TranslateAsync(
        string sourceText,
        string baseUrl,
        string? apiKey,
        string modelName,
        string targetLanguage,
        CancellationToken cancellationToken = default
    )
    {
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            throw new InvalidOperationException("Không có text để dịch.");
        }

        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            throw new InvalidOperationException("Thiếu URL cho Local AI (OpenAI-compatible).");
        }

        if (!Uri.TryCreate(baseUrl.Trim(), UriKind.Absolute, out var endpoint))
        {
            throw new InvalidOperationException("URL Local AI không hợp lệ.");
        }

        if (string.IsNullOrWhiteSpace(modelName))
        {
            throw new InvalidOperationException("Thiếu model cho Local AI.");
        }

        if (string.IsNullOrWhiteSpace(targetLanguage))
        {
            targetLanguage = "English";
        }

        var prompt =
            $"Translate the following text to natural {targetLanguage}. Return only the translation without additional explanation.\n\n"
            + sourceText;

        var clientOptions = new OpenAIClientOptions
        {
            Endpoint = endpoint
        };

        var credential = new ApiKeyCredential(string.IsNullOrWhiteSpace(apiKey) ? "local-ai-no-key" : apiKey.Trim());
        var chatClient = new ChatClient(modelName.Trim(), credential, clientOptions);
        var messages = new List<ChatMessage>
        {
            new UserChatMessage(prompt)
        };
        var completionOptions = new ChatCompletionOptions
        {
            Temperature = 0.2f
        };

        Exception? lastError = null;
        const int maxAttempts = 2; // first attempt + one retry

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                ClientResult<ChatCompletion> response = await chatClient.CompleteChatAsync(
                    messages,
                    completionOptions,
                    cancellationToken
                );

                var translated = string.Concat(
                        response.Value.Content
                            .Where(x => x.Kind == ChatMessageContentPartKind.Text)
                            .Select(x => x.Text)
                    )
                    .Trim();

                if (string.IsNullOrWhiteSpace(translated))
                {
                    throw new InvalidOperationException("Local AI không trả về nội dung dịch.");
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
                await Task.Delay(TimeSpan.FromMilliseconds(450), cancellationToken);
            }
            catch (Exception ex)
            {
                lastError = ex;
            }
        }

        throw lastError ?? new InvalidOperationException("Yêu cầu dịch Local AI thất bại.");
    }

    public void Dispose()
    {
        // No shared resource to dispose.
    }
}
