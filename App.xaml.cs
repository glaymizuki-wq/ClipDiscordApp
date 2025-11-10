using System.Windows;
using FormsApp = System.Windows.Forms;

namespace ClipDiscordApp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        // 例: WinForms の DoEvents を呼ぶ必要があるとき
        public void PumpWinFormsMessages()
        {
            FormsApp.Application.DoEvents();
        }

        // WPF の Application.Current を使う例
        public void Example()
        {
            var wpfApp = System.Windows.Application.Current;
        }
    }
}