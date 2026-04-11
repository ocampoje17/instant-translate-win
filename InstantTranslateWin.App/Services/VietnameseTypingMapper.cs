using InstantTranslateWin.App.Models;

namespace InstantTranslateWin.App.Services;

public static class VietnameseTypingMapper
{
    private enum TransformAction
    {
        None,
        ToneSac,
        ToneHuyen,
        ToneHoi,
        ToneNga,
        ToneNang,
        ShapeMu,
        ShapeHorn,
        ShapeBreve,
        ShapeDd,
        ShapeW
    }

    private enum VowelSeries
    {
        A,
        AW,
        AA,
        E,
        EE,
        I,
        O,
        OO,
        OW,
        U,
        UW,
        Y
    }

    private sealed record VowelMeta(VowelSeries Series, int Tone, bool IsUpper);

    private sealed record VowelPos(int Index, VowelSeries Series);

    private static readonly Dictionary<VowelSeries, (char[] Lower, char[] Upper)> VowelSeriesChars = BuildSeriesChars();
    private static readonly Dictionary<char, VowelMeta> VowelLookup = BuildVowelLookup();

    public static bool TryTransform(string currentText, string token, string typingStyle, out string transformedText)
    {
        var source = currentText ?? string.Empty;
        transformedText = source;

        if (string.IsNullOrEmpty(token) || token.Length != 1)
        {
            return false;
        }

        var input = token[0];
        if (!TryResolveAction(source, input, typingStyle, out var action))
        {
            transformedText = source + input;
            return true;
        }

        if (TryApplyAction(source, action, out var changed))
        {
            transformedText = changed;
            return true;
        }

        if (action == TransformAction.ShapeW && TryApplyStandaloneShapeW(source, input, out changed))
        {
            transformedText = changed;
            return true;
        }

        if (TryApplyRepeatedToneAsLiteral(source, action, out changed))
        {
            transformedText = changed + input;
            return true;
        }

        transformedText = source + input;
        return true;
    }

    private static bool TryResolveAction(string currentText, char input, string typingStyle, out TransformAction action)
    {
        action = TransformAction.None;
        var normalizedStyle = NormalizeStyle(typingStyle);
        var lower = char.ToLowerInvariant(input);

        if (string.Equals(normalizedStyle, QuickInputTypingOptions.VietnameseTypingStyleTelex, StringComparison.Ordinal))
        {
            action = lower switch
            {
                's' when CanApplyTelexTone(currentText, 's') => TransformAction.ToneSac,
                'f' when CanApplyTelexTone(currentText, 'f') => TransformAction.ToneHuyen,
                'r' when CanApplyTelexTone(currentText, 'r') => TransformAction.ToneHoi,
                'x' when CanApplyTelexTone(currentText, 'x') => TransformAction.ToneNga,
                'j' when CanApplyTelexTone(currentText, 'j') => TransformAction.ToneNang,
                'w' => TransformAction.ShapeW,
                'a' when CanConvertTelexShapeMu(currentText, VowelSeries.A) => TransformAction.ShapeMu,
                'e' when CanConvertTelexShapeMu(currentText, VowelSeries.E) => TransformAction.ShapeMu,
                'o' when CanConvertTelexShapeMu(currentText, VowelSeries.O) => TransformAction.ShapeMu,
                'd' when CanConvertLastDd(currentText) => TransformAction.ShapeDd,
                _ => TransformAction.None
            };

            return action != TransformAction.None;
        }

        if (string.Equals(normalizedStyle, QuickInputTypingOptions.VietnameseTypingStyleVni, StringComparison.Ordinal))
        {
            action = lower switch
            {
                '1' => TransformAction.ToneSac,
                '2' => TransformAction.ToneHuyen,
                '3' => TransformAction.ToneHoi,
                '4' => TransformAction.ToneNga,
                '5' => TransformAction.ToneNang,
                '6' => TransformAction.ShapeMu,
                '7' => TransformAction.ShapeHorn,
                '8' => TransformAction.ShapeBreve,
                '9' => TransformAction.ShapeDd,
                _ => TransformAction.None
            };

            return action != TransformAction.None;
        }

        // VIQR
        action = input switch
        {
            '\'' => TransformAction.ToneSac,
            '`' => TransformAction.ToneHuyen,
            '?' => TransformAction.ToneHoi,
            '~' => TransformAction.ToneNga,
            '.' => TransformAction.ToneNang,
            '^' => TransformAction.ShapeMu,
            '*' => TransformAction.ShapeHorn,
            '+' => TransformAction.ShapeHorn,
            '(' => TransformAction.ShapeBreve,
            '-' => TransformAction.ShapeDd,
            _ => TransformAction.None
        };

        return action != TransformAction.None;
    }

    private static bool TryApplyAction(string text, TransformAction action, out string transformedText)
    {
        transformedText = text;
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var wordStart = FindLastWordStart(text);
        if (wordStart < 0)
        {
            return false;
        }

        var word = text[wordStart..];
        return action switch
        {
            TransformAction.ToneSac => TryApplyTone(text, word, wordStart, tone: 1, out transformedText),
            TransformAction.ToneHuyen => TryApplyTone(text, word, wordStart, tone: 2, out transformedText),
            TransformAction.ToneHoi => TryApplyTone(text, word, wordStart, tone: 3, out transformedText),
            TransformAction.ToneNga => TryApplyTone(text, word, wordStart, tone: 4, out transformedText),
            TransformAction.ToneNang => TryApplyTone(text, word, wordStart, tone: 5, out transformedText),
            TransformAction.ShapeMu => TryApplyShape(text, word, wordStart, TransformAction.ShapeMu, out transformedText),
            TransformAction.ShapeHorn => TryApplyShape(text, word, wordStart, TransformAction.ShapeHorn, out transformedText),
            TransformAction.ShapeBreve => TryApplyShape(text, word, wordStart, TransformAction.ShapeBreve, out transformedText),
            TransformAction.ShapeDd => TryApplyDd(text, word, wordStart, out transformedText),
            TransformAction.ShapeW => TryApplyShapeW(text, word, wordStart, out transformedText),
            _ => false
        };
    }

    private static bool TryApplyTone(string fullText, string word, int wordStart, int tone, out string transformedText)
    {
        transformedText = fullText;
        var target = FindToneTargetIndex(word);
        if (target < 0)
        {
            return false;
        }

        var ch = word[target];
        if (!VowelLookup.TryGetValue(ch, out var meta))
        {
            return false;
        }

        var replacement = Compose(meta.Series, tone, meta.IsUpper);
        if (replacement == ch)
        {
            return false;
        }

        transformedText = ReplaceAt(fullText, wordStart + target, replacement);
        return true;
    }

    private static bool TryApplyRepeatedToneAsLiteral(string text, TransformAction action, out string transformedText)
    {
        transformedText = text;
        if (!TryGetToneValue(action, out var tone) || string.IsNullOrEmpty(text))
        {
            return false;
        }

        var wordStart = FindLastWordStart(text);
        if (wordStart < 0)
        {
            return false;
        }

        var word = text[wordStart..];
        var target = FindToneTargetIndex(word);
        if (target < 0)
        {
            return false;
        }

        var ch = word[target];
        if (!VowelLookup.TryGetValue(ch, out var meta) || meta.Tone != tone)
        {
            return false;
        }

        var replacement = Compose(meta.Series, 0, meta.IsUpper);
        if (replacement == ch)
        {
            return false;
        }

        transformedText = ReplaceAt(text, wordStart + target, replacement);
        return true;
    }

    private static bool TryGetToneValue(TransformAction action, out int tone)
    {
        tone = action switch
        {
            TransformAction.ToneSac => 1,
            TransformAction.ToneHuyen => 2,
            TransformAction.ToneHoi => 3,
            TransformAction.ToneNga => 4,
            TransformAction.ToneNang => 5,
            _ => 0
        };

        return tone != 0;
    }

    private static bool TryApplyShape(string fullText, string word, int wordStart, TransformAction shape, out string transformedText)
    {
        transformedText = fullText;

        for (var i = word.Length - 1; i >= 0; i--)
        {
            if (!TryConvertVowel(word[i], shape, out var replacement))
            {
                continue;
            }

            transformedText = ReplaceAt(fullText, wordStart + i, replacement);
            return true;
        }

        return false;
    }

    private static bool TryApplyShapeW(string fullText, string word, int wordStart, out string transformedText)
    {
        transformedText = fullText;
        if (string.IsNullOrEmpty(word))
        {
            return false;
        }

        if (TryApplyUoHornPair(fullText, word, wordStart, out transformedText))
        {
            return true;
        }

        for (var i = word.Length - 1; i >= 0; i--)
        {
            if (TryConvertVowel(word[i], TransformAction.ShapeBreve, out var breveReplacement))
            {
                transformedText = ReplaceAt(fullText, wordStart + i, breveReplacement);
                return true;
            }

            if (TryConvertVowel(word[i], TransformAction.ShapeHorn, out var hornReplacement))
            {
                transformedText = ReplaceAt(fullText, wordStart + i, hornReplacement);
                return true;
            }
        }

        return false;
    }

    private static bool TryApplyStandaloneShapeW(string text, char input, out string transformedText)
    {
        transformedText = text;
        if (char.ToLowerInvariant(input) != 'w')
        {
            return false;
        }

        transformedText = text + Compose(VowelSeries.UW, 0, char.IsUpper(input));
        return true;
    }

    private static bool TryApplyUoHornPair(string fullText, string word, int wordStart, out string transformedText)
    {
        transformedText = fullText;
        if (word.Length < 2)
        {
            return false;
        }

        var oIndex = word.Length - 1;
        var uIndex = word.Length - 2;

        if (!VowelLookup.TryGetValue(word[uIndex], out var uMeta) || !VowelLookup.TryGetValue(word[oIndex], out var oMeta))
        {
            return false;
        }

        if (uMeta.Series != VowelSeries.U || oMeta.Series != VowelSeries.O)
        {
            return false;
        }

        transformedText = ReplaceAt(fullText, wordStart + uIndex, Compose(VowelSeries.UW, uMeta.Tone, uMeta.IsUpper));
        transformedText = ReplaceAt(transformedText, wordStart + oIndex, Compose(VowelSeries.OW, oMeta.Tone, oMeta.IsUpper));
        return true;
    }

    private static bool TryApplyDd(string fullText, string word, int wordStart, out string transformedText)
    {
        transformedText = fullText;
        for (var i = word.Length - 1; i >= 0; i--)
        {
            var ch = word[i];
            char replacement;
            if (ch == 'd')
            {
                replacement = 'đ';
            }
            else if (ch == 'D')
            {
                replacement = 'Đ';
            }
            else
            {
                continue;
            }

            transformedText = ReplaceAt(fullText, wordStart + i, replacement);
            return true;
        }

        return false;
    }

    private static int FindToneTargetIndex(string word)
    {
        var vowels = new List<VowelPos>();
        for (var i = 0; i < word.Length; i++)
        {
            if (!VowelLookup.TryGetValue(word[i], out var meta))
            {
                continue;
            }

            vowels.Add(new VowelPos(i, meta.Series));
        }

        if (vowels.Count == 0)
        {
            return -1;
        }

        // "qu" and "gi" leading sequences: the second character often behaves like a consonant.
        if (vowels.Count > 1)
        {
            if (word.Length >= 2 &&
                (word[0] == 'q' || word[0] == 'Q'))
            {
                vowels.RemoveAll(v => v.Index == 1 && (v.Series == VowelSeries.U || v.Series == VowelSeries.UW));
            }

            if (word.Length >= 2 &&
                (word[0] == 'g' || word[0] == 'G'))
            {
                vowels.RemoveAll(v => v.Index == 1 && v.Series == VowelSeries.I);
            }
        }

        if (vowels.Count == 0)
        {
            return -1;
        }

        var shaped = vowels.Where(v => IsShapedSeries(v.Series)).ToList();
        if (shaped.Count == 1)
        {
            return shaped[0].Index;
        }

        if (shaped.Count > 1)
        {
            return shaped[^1].Index;
        }

        if (vowels.Count == 1)
        {
            return vowels[0].Index;
        }

        if (vowels.Count == 2)
        {
            return vowels[1].Index == word.Length - 1
                ? vowels[0].Index
                : vowels[1].Index;
        }

        return vowels[1].Index;
    }

    private static bool TryConvertVowel(char source, TransformAction shape, out char converted)
    {
        converted = source;
        if (!VowelLookup.TryGetValue(source, out var meta))
        {
            return false;
        }

        VowelSeries target = VowelSeries.A;
        var hasTarget = true;
        switch (shape)
        {
            case TransformAction.ShapeMu:
                target = meta.Series switch
                {
                    VowelSeries.A => VowelSeries.AA,
                    VowelSeries.E => VowelSeries.EE,
                    VowelSeries.O => VowelSeries.OO,
                    _ => target
                };
                hasTarget = meta.Series is VowelSeries.A or VowelSeries.E or VowelSeries.O;
                break;
            case TransformAction.ShapeHorn:
                target = meta.Series switch
                {
                    VowelSeries.O => VowelSeries.OW,
                    VowelSeries.U => VowelSeries.UW,
                    _ => target
                };
                hasTarget = meta.Series is VowelSeries.O or VowelSeries.U;
                break;
            case TransformAction.ShapeBreve:
                target = meta.Series == VowelSeries.A
                    ? VowelSeries.AW
                    : target;
                hasTarget = meta.Series == VowelSeries.A;
                break;
            default:
                hasTarget = false;
                break;
        }

        if (!hasTarget)
        {
            return false;
        }

        converted = Compose(target, meta.Tone, meta.IsUpper);
        return converted != source;
    }

    private static bool CanConvertTelexShapeMu(string text, VowelSeries targetSeries)
    {
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var wordStart = FindLastWordStart(text);
        if (wordStart < 0)
        {
            return false;
        }

        var word = text[wordStart..];
        return CanConvertTrailingBaseVowel(word, targetSeries) || CanConvertBaseVowelBeforeTerminalSemiVowel(word, targetSeries);
    }

    private static bool CanApplyTelexTone(string text, char toneKey)
    {
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var wordStart = FindLastWordStart(text);
        if (wordStart < 0)
        {
            return false;
        }

        var lastChar = text[^1];
        return !char.Equals(char.ToLowerInvariant(lastChar), toneKey);
    }

    private static bool CanConvertTrailingBaseVowel(string word, VowelSeries targetSeries)
    {
        if (string.IsNullOrEmpty(word))
        {
            return false;
        }

        var lastChar = word[^1];
        if (!VowelLookup.TryGetValue(lastChar, out var meta))
        {
            return false;
        }

        var matchesTargetSeries = targetSeries switch
        {
            VowelSeries.A => meta.Series == VowelSeries.A,
            VowelSeries.E => meta.Series == VowelSeries.E,
            VowelSeries.O => meta.Series == VowelSeries.O,
            _ => false
        };

        if (!matchesTargetSeries)
        {
            return false;
        }

        for (var i = 0; i < word.Length - 1; i++)
        {
            if (!VowelLookup.TryGetValue(word[i], out var previousMeta))
            {
                continue;
            }

            if (previousMeta.Series == targetSeries)
            {
                return false;
            }
        }

        return true;
    }

    private static bool CanConvertBaseVowelBeforeTerminalSemiVowel(string word, VowelSeries targetSeries)
    {
        if (word.Length < 2)
        {
            return false;
        }

        var lastChar = word[^1];
        if (!VowelLookup.TryGetValue(lastChar, out var meta))
        {
            return false;
        }

        if (meta.Series is not (VowelSeries.I or VowelSeries.U or VowelSeries.Y))
        {
            return false;
        }

        var candidateChar = word[^2];
        if (!VowelLookup.TryGetValue(candidateChar, out var candidateMeta))
        {
            return false;
        }

        var matchesTargetSeries = targetSeries switch
        {
            VowelSeries.A => candidateMeta.Series == VowelSeries.A,
            VowelSeries.E => candidateMeta.Series == VowelSeries.E,
            VowelSeries.O => candidateMeta.Series == VowelSeries.O,
            _ => false
        };

        if (!matchesTargetSeries)
        {
            return false;
        }

        for (var i = 0; i < word.Length - 2; i++)
        {
            if (!VowelLookup.TryGetValue(word[i], out var previousMeta))
            {
                continue;
            }

            if (previousMeta.Series == targetSeries)
            {
                return false;
            }
        }

        return true;
    }

    private static bool CanConvertLastDd(string text)
    {
        return !string.IsNullOrEmpty(text) && (text[^1] == 'd' || text[^1] == 'D');
    }

    private static int FindLastWordStart(string text)
    {
        var i = text.Length - 1;
        if (i < 0 || !IsWordChar(text[i]))
        {
            return -1;
        }

        while (i >= 0 && IsWordChar(text[i]))
        {
            i--;
        }

        return i + 1;
    }

    private static bool IsWordChar(char ch)
    {
        return char.IsLetter(ch);
    }

    private static bool IsShapedSeries(VowelSeries series)
    {
        return series is VowelSeries.AW or VowelSeries.AA or VowelSeries.EE or VowelSeries.OO or VowelSeries.OW or VowelSeries.UW;
    }

    private static string ReplaceAt(string source, int index, char replacement)
    {
        return source[..index] + replacement + source[(index + 1)..];
    }

    private static char Compose(VowelSeries series, int tone, bool isUpper)
    {
        var table = VowelSeriesChars[series];
        var safeTone = Math.Clamp(tone, 0, 5);
        return isUpper ? table.Upper[safeTone] : table.Lower[safeTone];
    }

    private static string NormalizeStyle(string style)
    {
        if (string.Equals(style, QuickInputTypingOptions.VietnameseTypingStyleViqr, StringComparison.OrdinalIgnoreCase))
        {
            return QuickInputTypingOptions.VietnameseTypingStyleViqr;
        }

        if (string.Equals(style, QuickInputTypingOptions.VietnameseTypingStyleVni, StringComparison.OrdinalIgnoreCase))
        {
            return QuickInputTypingOptions.VietnameseTypingStyleVni;
        }

        return QuickInputTypingOptions.VietnameseTypingStyleTelex;
    }

    private static Dictionary<char, VowelMeta> BuildVowelLookup()
    {
        var result = new Dictionary<char, VowelMeta>();
        foreach (var pair in VowelSeriesChars)
        {
            var series = pair.Key;
            var lower = pair.Value.Lower;
            var upper = pair.Value.Upper;
            for (var tone = 0; tone < 6; tone++)
            {
                result[lower[tone]] = new VowelMeta(series, tone, IsUpper: false);
                result[upper[tone]] = new VowelMeta(series, tone, IsUpper: true);
            }
        }

        return result;
    }

    private static Dictionary<VowelSeries, (char[] Lower, char[] Upper)> BuildSeriesChars()
    {
        return new Dictionary<VowelSeries, (char[] Lower, char[] Upper)>
        {
            [VowelSeries.A] = ("aáàảãạ".ToCharArray(), "AÁÀẢÃẠ".ToCharArray()),
            [VowelSeries.AW] = ("ăắằẳẵặ".ToCharArray(), "ĂẮẰẲẴẶ".ToCharArray()),
            [VowelSeries.AA] = ("âấầẩẫậ".ToCharArray(), "ÂẤẦẨẪẬ".ToCharArray()),
            [VowelSeries.E] = ("eéèẻẽẹ".ToCharArray(), "EÉÈẺẼẸ".ToCharArray()),
            [VowelSeries.EE] = ("êếềểễệ".ToCharArray(), "ÊẾỀỂỄỆ".ToCharArray()),
            [VowelSeries.I] = ("iíìỉĩị".ToCharArray(), "IÍÌỈĨỊ".ToCharArray()),
            [VowelSeries.O] = ("oóòỏõọ".ToCharArray(), "OÓÒỎÕỌ".ToCharArray()),
            [VowelSeries.OO] = ("ôốồổỗộ".ToCharArray(), "ÔỐỒỔỖỘ".ToCharArray()),
            [VowelSeries.OW] = ("ơớờởỡợ".ToCharArray(), "ƠỚỜỞỠỢ".ToCharArray()),
            [VowelSeries.U] = ("uúùủũụ".ToCharArray(), "UÚÙỦŨỤ".ToCharArray()),
            [VowelSeries.UW] = ("ưứừửữự".ToCharArray(), "ƯỨỪỬỮỰ".ToCharArray()),
            [VowelSeries.Y] = ("yýỳỷỹỵ".ToCharArray(), "YÝỲỶỸỴ".ToCharArray())
        };
    }
}
