using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using InstantTranslateWin.App.Services;

namespace InstantTranslateWin.App.Controls;

public partial class SupportInfoCard : UserControl
{
    public SupportInfoCard()
    {
        InitializeComponent();
        AppVersionTextBlock.Text = $"Phiên bản: {ResolveAppVersion()}";
    }

    private void SupportLink_OnRequestNavigate(object sender, RequestNavigateEventArgs e)
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
            ErrorFileLogger.LogException("SupportInfoCard.SupportLink_OnRequestNavigate", ex);
            MessageBox.Show(
                $"Không mở được liên kết: {ex.Message}",
                "Instant Translate",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            );
        }
    }

    private static string ResolveAppVersion()
    {
        try
        {
            var assembly = typeof(App).Assembly;
            var informationalVersion = assembly
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                ?.InformationalVersion;

            if (!string.IsNullOrWhiteSpace(informationalVersion))
            {
                // Keep semantic version only; hide build metadata/hash.
                var plusIndex = informationalVersion.IndexOf('+');
                return plusIndex > 0
                    ? informationalVersion[..plusIndex]
                    : informationalVersion;
            }

            var fileVersion = assembly
                .GetCustomAttribute<AssemblyFileVersionAttribute>()
                ?.Version;

            if (!string.IsNullOrWhiteSpace(fileVersion))
            {
                return fileVersion;
            }

            return assembly.GetName().Version?.ToString() ?? "N/A";
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("SupportInfoCard.ResolveAppVersion", ex);
            return "N/A";
        }
    }
}
