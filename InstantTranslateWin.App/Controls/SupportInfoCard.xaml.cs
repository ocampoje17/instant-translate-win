using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace InstantTranslateWin.App.Controls;

public partial class SupportInfoCard : UserControl
{
    public SupportInfoCard()
    {
        InitializeComponent();
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
            MessageBox.Show(
                $"Không mở được liên kết: {ex.Message}",
                "Instant Translate",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            );
        }
    }
}
