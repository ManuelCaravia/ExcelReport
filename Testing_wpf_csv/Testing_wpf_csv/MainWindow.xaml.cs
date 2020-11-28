using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Testing_wpf_csv.Control;


namespace Testing_wpf_csv
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string file_path = @"";
        Controller controller;
        
        public MainWindow()
        {
            InitializeComponent();            
        }

        private void New_location_btn_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            string[] csv_files;//addresses of all csv files found
            DialogResult result = folderBrowser.ShowDialog();
            file_path = folderBrowser.SelectedPath;
            
            csv_files = Directory.GetFiles(file_path, "*.csv");
            
            controller = new Controller(csv_files);
            controller.Process_files();
        }
    }
}
