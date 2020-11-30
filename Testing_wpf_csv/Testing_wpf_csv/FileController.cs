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
        private int amount_of_records;

        public string Video_location { get => video_data_location; set => video_data_location = value; }        
        public string Exported_file_path { get => exported_file_path; set => exported_file_path = value; }

        public FileController(string csv_file_path)
        {
            video_data_location = csv_file_path;
            db = new DataHandler();            
            exported_file_path = Path.ChangeExtension(video_data_location, "xlsx");
            db.CreateNewFile(exported_file_path);
            db.Open(exported_file_path, 1);
            db.RenameSheet("Summary");
            db.CreateNewSheet();
            db.SelectSheet(2);
            db.RenameSheet("raw data");
            db.SaveWorkBookProgress();
            db.Close();            
        }        
        

       /* calls DataHandler object to create a new file
         * then uses it to initialize columns and save summary 
        */
        public void WriteSummary()
        {
            string temp_str;
            
            db.Open(exported_file_path, "Summary");
            //init title columns 
            db.WriteToCell(4, 2, "Shoot height statistics");
            db.WriteToCell(6, 3, "Number of shoots which have reached this shoot height (m)");
            db.WriteToCell(7, 2, "Video");
            db.WriteToCell(8, 2, Exported_file_path);


            //Writing amount of shoots to reach to a certain height and the total percentage            

            //Over 0.6 m 
            db.WriteToCell(7, 3, "0.6");            
            temp_str = "=COUNTIF('raw data'!A2:A" + amount_of_records.ToString() + ",\" > 0.6\")";
            db.WriteFormulaToCell(8,3, temp_str);
            db.WriteFormulaToCell(9, 3, "=C8/B9");
            //Over 0.8
            db.WriteToCell(7, 4, "0.8");
            temp_str = "=COUNTIF('raw data'!A2:A" + amount_of_records.ToString() + ",\" > 0.8\")";
            db.WriteFormulaToCell(8, 4, temp_str);
            db.WriteFormulaToCell(9, 4, "=D8/B9");
            //Over 1m 
            db.WriteToCell(7, 5, "1");
            temp_str = "=COUNTIF('raw data'!A2:A" + amount_of_records.ToString() + ",\" > 1\")";
            db.WriteFormulaToCell(8, 5, temp_str);
            db.WriteFormulaToCell(9, 5, "=E8/B9");
            //over 1.2m
            db.WriteToCell(7, 6, "1.2");
            temp_str = "=COUNTIF('raw data'!A2:A" + amount_of_records.ToString() + ",\" > 1.2\")";
            db.WriteFormulaToCell(8, 6, temp_str);
            db.WriteFormulaToCell(9, 6, "=F8/B9");
            //over 1.4m
            db.WriteToCell(7, 7, "1.4");
            temp_str = "=COUNTIF('raw data'!A2:A" + amount_of_records.ToString() + ",\" > 1.4\")";
            db.WriteFormulaToCell(8, 7, temp_str);
            db.WriteFormulaToCell(9, 7, "=G8/B9");

            db.WriteToCell(9, 1, "Total");
            db.WriteToCell(9, 2, amount_of_records);            
            db.SaveWorkBookProgress();
            db.Close();
        }       
        public void MergeRawData()
        {
            db.Open(exported_file_path,"raw data");
            List<string> csv_lines = db.GetRawDataList(video_data_location);
            int line_count = 0;
            double temp_num;
            db.SelectSheet(2);
            db.WriteToCell(1, 1, "time");
            db.WriteToCell(1, 2, "average canopy height");
            db.WriteToCell(1, 3, "latitud");
            db.WriteToCell(1, 4, "longitud");


            foreach (var line in csv_lines)
            {
                if (line_count > 0)//if not title row
                {
                    string[] entries = line.Split(',');
                    //time of record
                    temp_num = Double.Parse(entries[0]);
                    db.WriteToCell(line_count+1, 1, temp_num);

                    //average height
                    temp_num = Double.Parse(entries[1]);
                    db.WriteToCell(line_count + 1, 2, temp_num);

                    //latitud
                    temp_num = Double.Parse(entries[2]);
                    db.WriteToCell(line_count + 1, 3, temp_num);
                    //longitud
                    temp_num = Double.Parse(entries[3]);
                    db.WriteToCell(line_count + 1, 4, temp_num);
                }                                               
                line_count++;
            }
            amount_of_records = line_count;
            db.SaveWorkBookProgress();
            db.Close();
        }
        public void DrawGraph_Style()
        {
            db.Open(exported_file_path, "Summary");
            //Styling
            db.StyleExcelFile(amount_of_records);
            db.SaveWorkBookProgress();
            db.Close();
        }
        //anything that needs to be done at the end, in this case select which sheet is desplayed when used open doc
        public void FinishUp()
        {
            db.Open(exported_file_path, "Summary");            
            db.SelectOpenSheet();
            db.SaveWorkBookProgress();
        }
        public void ProcessFile()
        {            
            MergeRawData();
            WriteSummary();
            DrawGraph_Style();
            FinishUp();
        }
    }
            
}
