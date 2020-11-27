using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Testing_wpf_csv.Models;

namespace Testing_wpf_csv.Control
{
    class Controller
    {
        private string video_data_location = @"D:\desk\canopy-height-report-Template-20201124.xlsx";
        private DataHandler db;
        private int[] sizes_count = new int[5];//0= under 0.6 m, 2 = under 0.8, 2 = under 1m, 3 = under 1.2, 4 = under 1.4,
        private int min_height;//shortest height found 
        private int average_height;//average of all samples
        private int max_height;//highest height of all samples 
        private ObservableCollection<RawRecord> raw_data;        

        public string Video_location { get => video_data_location; set => video_data_location = value; }
        public int[] Sizes_count { get => sizes_count; set => sizes_count = value; }
        public int Min_height { get => min_height; set => min_height = value; }
        public int Average_height { get => average_height; set => average_height = value; }
        public int Max_height { get => max_height; set => max_height = value; }
        public ObservableCollection<RawRecord> Raw_data { get => raw_data; set => raw_data = value; }
        public void Load_raw_data(string filepath)
        {
            db = new DataHandler(video_data_location, 2);
            raw_data = new ObservableCollection<RawRecord>();
            int num_rows = db.GetRowCount();

            for (int row = 2; row <= num_rows; row++ )
            {
                double time= db.ReadCell(row, 1);
                double average_shoot_height = db.ReadCell(row, 2);
                double latitud = db.ReadCell(row, 3);
                double longitud = db.ReadCell(row, 4);
                raw_data.Add(new RawRecord(time, average_shoot_height, latitud, longitud));                
            }
        } 
    }
}
