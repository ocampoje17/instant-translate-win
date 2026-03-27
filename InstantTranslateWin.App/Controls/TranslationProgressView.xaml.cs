using System.Windows.Controls;

namespace InstantTranslateWin.App.Controls;

public partial class TranslationProgressView : UserControl
{
    public TranslationProgressView()
    {
        InitializeComponent();
    }

    public string HeaderText
    {
        get => HeaderTextBlock.Text;
        set => HeaderTextBlock.Text = value;
    }

    public void SetProgress(string status, double progressPercent)
    {
        StatusTextBlock.Text = status;
        ProgressIndicator.IsIndeterminate = false;
        ProgressIndicator.Value = Math.Clamp(progressPercent, 0, 100);
    }

    public void SetIdle(string status)
    {
        SetProgress(status, 0);
    }
}
