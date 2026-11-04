using System.Windows;
using ToolVba.App.ViewModels;

namespace ToolVba.App;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel(this);
    }
}
