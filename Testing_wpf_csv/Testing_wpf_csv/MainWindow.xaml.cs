using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
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
        static string file_path = @"D:\desk\canopy-height-report-Template-20201124.xlsx";
        Controller controller;
        public MainWindow()
        {
            InitializeComponent();
            controller = new Controller();
            controller.Load_raw_data(file_path);
            raw_listview.ItemsSource = controller.Raw_data;
        }

        private void New_file_button_Click(object sender, RoutedEventArgs e)
        {            
            testing.Text = controller.Raw_data.Count.ToString();            
        }

        private void Write_button_Click(object sender, RoutedEventArgs e)
        {
            int rows_to_right = Int32.Parse(row_picker.Text);
            int columns_down = Int32.Parse(column_picker.Text);                      
        }
    }
}
