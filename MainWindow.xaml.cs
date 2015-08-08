using System;
using System.Windows;
using System.Windows.Forms;

namespace FinAnParser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static System.Windows.Controls.TextBox tbResult;
        public static string File { get; set; }
        public static string Folder { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            tbResult = TextBoxRESULTINFO;
        }

        private void ButtonFROM_Click(object sender, RoutedEventArgs e)
        {

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".txt";
            dlg.Filter = "HTM (*.htm)|*.htm|HTML (*.html)|*.html";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                File = dlg.FileName;
                TextBoxFROM.Text = File;
            }

            if (File != null & Folder != null)
            {
                ButtonGO.IsEnabled = true;
            }

            TextBoxRESULTINFO.Text = "";
        }

        private void ButtonWHERE_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Folder = dlg.SelectedPath;
                TextBoxWHERE.Text = Folder;
            }
            else
            {
                // This prevents a crash when you close out of the window with nothing
            }

            if (File != null & Folder != null)
            {
                ButtonGO.IsEnabled = true;
            }

            TextBoxRESULTINFO.Text = "";
        }

        private void ButtonGO_Click(object sender, RoutedEventArgs e)
        {
            new HtmlParser().Parse();        
        }
    }
}
