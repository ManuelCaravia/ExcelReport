using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private Controller controller;
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Click_choose_folder(object sender, RoutedEventArgs e)
        {
            btn_location.IsEnabled = false;            
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            string[] csv_files;//path of csv files found
            DialogResult result = folderBrowser.ShowDialog();
            if (string.IsNullOrEmpty(folderBrowser.SelectedPath))
            {
                info1.Text = "No Folder Selected";
            }
            else
            {
                info1.Text = "Processing Path Selected";
                file_path = folderBrowser.SelectedPath;

                csv_files = Directory.GetFiles(file_path, "*.csv");
                if (csv_files.Length == 0)
                {
                    info2.Text = "No CSV Files Found";
                    info1.Text = "";
                    btn_process_folder.IsEnabled = false;
                    clearBtn.Visibility= Visibility.Visible ;
                }
                else
                {
                    info1.Text = "Processing";
                    info2.Text = csv_files.Length + " CSV Files Found";
                    controller = new Controller(csv_files);
                    btn_location.IsEnabled = true;
                    btn_process_folder.IsEnabled = false;
                    info1.Text = "Finished Processing";
                    info2.Text = "Select a Different Folder";
                }
                
            }  
        }


        private void Click_process_folder(object sender, RoutedEventArgs e)
        {
            btn_location.IsEnabled = false;
            btn_process_folder.IsEnabled = false;
            controller.ProcessFiles();
            btn_location.IsEnabled = true;
            btn_process_folder.IsEnabled = true;
            info1.Text = "Finished Processing";
            info2.Text = "";
            clearBtn.Visibility = Visibility.Visible;
        }

        private void ClearBtn_Click(object sender, RoutedEventArgs e)
        {
            info1.Text = "";
            info2.Text = "";
            info3.Text = "Hint: Specify a folder that that contains raw data CSV files";
            btn_location.IsEnabled = true;
            btn_process_folder.IsEnabled = false;
            clearBtn.Visibility = Visibility.Hidden;
        }
    }
}
