using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Testing_wpf_csv.Models;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Testing_wpf_csv
{
    class DataHandler
    {
        private string raw_data_file_location= @"";        

        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        


        public string Raw_data_file_location { get => raw_data_file_location; set => raw_data_file_location = value; }

        public DataHandler(string path, int sheet)
        {            
            raw_data_file_location = path;
            wb = excel.Workbooks.Open(raw_data_file_location);
            ws = wb.Worksheets[sheet];                    
        }

        public double ReadCell(int cells_down, int cells_to_right)
        {
            if (ws.Cells[cells_down, cells_to_right].Value2 != null)
            {
                return ws.Cells[cells_down, cells_to_right].Value2;
            }
            else return -1.0;
        }

        public void WriteToCell(int cells_down, int cells_to_right, string s)
        {
            ws.Cells[cells_down, cells_to_right].Value2 = s;
            wb.Save();
        }
        public int GetRowCount()
        {
            int row_count = 0;
            while (ws.Cells[(row_count+1), 1].Value2 != null)// +1 since excel doesnt have row 0, starts from 1
            {
                row_count=row_count + 30;//we increment by 30 first to reduce amount of checks
            }
            row_count = row_count - 30;
            while (ws.Cells[(row_count + 1), 1].Value2 != null)
            {
                row_count++;
            }
            return row_count;
        }
        public void CreateNewFile(string export_excel_path)
        {
            wb = excel.Workbooks.Add();
            wb.SaveAs(export_excel_path);
            ws = wb.Worksheets[1];            
        }
        public void Close()
        {
            wb.Close();
        }
        
    }
}
