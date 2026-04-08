using System.Windows;

namespace InventorAutoSave.Setup
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            var win = new SetupWindow();
            win.Show();
        }
    }
}
