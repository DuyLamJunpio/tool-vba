using Velopack;

namespace ToolVba.App;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        // Must run first: handles install/update/uninstall hooks and exits early when Velopack asks it to.
        VelopackApp.Build().Run();

        var app = new App();
        app.InitializeComponent();
        app.Run();
    }
}
