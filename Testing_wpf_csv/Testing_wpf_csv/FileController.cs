using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Testing_wpf_csv.Models;

namespace Testing_wpf_csv.Control
{
    class FileController
    {
        private string video_data_location = @"";
        private string exported_file_path = @"";
        private DataHandler db;
        private int[] sizes_count = new int[5];//0= under 0.6 m, 1 = under 0.8, 2 = under 1m, 3 = under 1.2, 4 = under 1.4,
        private double[] sizes_count_percentage = new double[5];//0= under 0.6 m, 1 = under 0.8, 2 = under 1m, 3 = under 1.2, 4 = under 1.4,
        private double min_height;//shortest height found 
        private double average_height;//average of all samples
        private double max_height;//highest height of all samples 
        private List<RawRecord> raw_data;
        private bool dataFetched,dataProcessed = false ;
        private int amount_of_records;

        public string Video_location { get => video_data_location; set => video_data_location = value; }
        public int[] Sizes_count { get => sizes_count; set => sizes_count = value; }
        public double Min_height { get => min_height; set => min_height = value; }
        public double Average_height { get => average_height; set => average_height = value; }
        public double Max_height { get => max_height; set => max_height = value; }
        public List<RawRecord> Raw_data { get => raw_data; set => raw_data = value; }
        public string Exported_file_path { get => exported_file_path; set => exported_file_path = value; }

        public FileController(string full_file_path)
        {
            video_data_location = full_file_path;
            db = new DataHandler(video_data_location, 1);
            raw_data = new List<RawRecord>();
            exported_file_path = Path.ChangeExtension(video_data_location, "xlsx");
        }
        public void Load_raw_data()
        {            
            int num_rows = db.GetRowCount();

            for (int row = 2; row <= num_rows; row++ )
            {
                double time= db.ReadCell(row, 1);
                double average_shoot_height = db.ReadCell(row, 2);
                double latitud = db.ReadCell(row, 3);
                double longitud = db.ReadCell(row, 4);
                raw_data.Add(new RawRecord(time, average_shoot_height, latitud, longitud));                
            }
            dataFetched = true;
        }
        public void ProcessData()
        {
            if(!dataFetched)
            {
                Load_raw_data();
            }
            double height_sum = 0; //will be used to calc average height of all shoots
            amount_of_records = raw_data.Count;


            //add values to variables  
            min_height = 1000;
            max_height = 0;
            for(int height = 0;height<5; height++)
            {
                sizes_count[height] = 0;//                
            }
            foreach(var record in raw_data)
            {
                //if average_shoot_height larger than 0.6
                if (record.Average_shoot_height > 0.6)
                {
                    sizes_count[0]++;
                }
                //if average_shoot_height larger than 0.8
                if (record.Average_shoot_height > 0.8)
                {
                    sizes_count[1]++;
                }
                //if average_shoot_height larger than 1.0
                if (record.Average_shoot_height > 1)
                {
                    sizes_count[2]++;
                }
                //if average_shoot_height larger than 1.2
                if (record.Average_shoot_height > 1.2)
                {
                    sizes_count[3]++;
                }
                //if average_shoot_height larger than 1.4
                if (record.Average_shoot_height > 1.4)
                {
                    sizes_count[4]++;
                }
                //end of section that calculated number of shoots larger than a certain height 


                //now check for min and max height values 
                if(record.Average_shoot_height<min_height)
                {
                    min_height = record.Average_shoot_height;
                }
                if(record.Average_shoot_height>max_height)
                {
                    max_height = record.Average_shoot_height;
                }
                height_sum = height_sum + record.Average_shoot_height;
            }
            average_height = height_sum / amount_of_records;//calc average height of all shoots

            //calc percentage of shoot heights
            for (int height = 0; height < 5; height++)
            {
                sizes_count_percentage[height] = ((double)sizes_count[height]/ (double)amount_of_records)*100.0;//                
            }
            dataProcessed = true;
        }      
        
        /* calls DataHandler object to create a new file
         * then uses it to initialize columns and save summary 
        */
        public void SaveSummary()
        {
            if (!dataProcessed)
            {
                ProcessData();
            }

            db.CreateNewFile(@exported_file_path);
            //init title columns 
            db.WriteToCell(4, 2, "Shoot height statistics");
            db.WriteToCell(6, 3, "Number of shoots which have reached this shoot height (m)");
            db.WriteToCell(7, 2, "Video");
            db.WriteToCell(8, 2, Exported_file_path);

            //Writing amount of shoots to reach to a certain height and the total percentage            
            db.WriteToCell(7,3,"0.6");
            db.WriteToCell(8, 3, sizes_count[0].ToString());
            db.WriteToCell(9, 3, sizes_count_percentage[0].ToString());

            db.WriteToCell(7, 4, "0.8");
            db.WriteToCell(8, 4, sizes_count[1].ToString());
            db.WriteToCell(9, 4, sizes_count_percentage[1].ToString());

            db.WriteToCell(7, 5, "1");
            db.WriteToCell(8, 5, sizes_count[2].ToString());
            db.WriteToCell(9, 5, sizes_count_percentage[2].ToString());

            db.WriteToCell(7, 6, "1.2");
            db.WriteToCell(8, 6, sizes_count[3].ToString());
            db.WriteToCell(9, 6, sizes_count_percentage[3].ToString());

            db.WriteToCell(7, 7, "1.4");
            db.WriteToCell(8, 7, sizes_count[4].ToString());
            db.WriteToCell(9, 7, sizes_count_percentage[4].ToString());

            db.WriteToCell(9, 1, "Total");
            db.WriteToCell(9, 2, amount_of_records.ToString());

            db.Close();
        }
    }
}
