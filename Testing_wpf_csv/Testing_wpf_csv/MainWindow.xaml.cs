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
        Controller controller;
        
        public MainWindow()
        {
            InitializeComponent();            
        }

        private void New_location_btn_Click(object sender, RoutedEventArgs e)
        {
            new_location_btn.IsEnabled = false;            
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            string[] csv_files;//addresses of all csv files found
            DialogResult result = folderBrowser.ShowDialog();            
            file_path = folderBrowser.SelectedPath;
            
            csv_files = Directory.GetFiles(file_path, "*.csv");
            
            controller = new Controller(csv_files);
            new_location_btn.IsEnabled = true;
            process_btn.IsEnabled = true;
            info1.Text = "";
        }


        private void Process_btn_Click(object sender, RoutedEventArgs e)
        {
            new_location_btn.IsEnabled = false;
            process_btn.IsEnabled = false;
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += BackgroundWorker1_DoWork;
            worker.ProgressChanged += BackgroundWorker1_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.RunWorkerAsync();
        }
        private void BackgroundWorker1_DoWork(object sender,
            DoWorkEventArgs e)
        {
                        
            controller.ProcessFiles(sender);
        }


        // This event handler updates the progress bar.
        private void BackgroundWorker1_ProgressChanged(object sender,
            ProgressChangedEventArgs e)
        {
            //this.progressBar1.Value = e.ProgressPercentage;
            info1.Text = e.ProgressPercentage.ToString() + " files processed";
            info2.Text = "Working on file: " + e.UserState.ToString();
        }
        void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {            
            System.Windows.Forms.MessageBox.Show("Done " );
            info1.Text = "";
            info2.Text = "";
            new_location_btn.IsEnabled = true;            
        }
    }
}
