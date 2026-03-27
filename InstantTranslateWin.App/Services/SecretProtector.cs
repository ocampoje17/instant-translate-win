using System.Security.Cryptography;
using System.Text;

namespace InstantTranslateWin.App.Services;

public static class SecretProtector
{
    public static string? Protect(string? plainText)
    {
        if (string.IsNullOrWhiteSpace(plainText))
        {
            return null;
        }

        var bytes = Encoding.UTF8.GetBytes(plainText);
        var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(protectedBytes);
    }

    public static string Unprotect(string? cipherText)
    {
        if (string.IsNullOrWhiteSpace(cipherText))
        {
            return string.Empty;
        }

        try
        {
            var bytes = Convert.FromBase64String(cipherText);
            var plainBytes = ProtectedData.Unprotect(bytes, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plainBytes);
        }
        catch
        {
            return string.Empty;
        }
    }
}
