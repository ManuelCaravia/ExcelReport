using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Testing_wpf_csv.Models;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Testing_wpf_csv
{
    class DataHandler
    {
        private string raw_data_file_location= @"D:\desk\canopy-height-report-Template-20201124.xlsx";
        

        _Application excel = new _Excel.Application();

        Workbook wb;

        Worksheet ws;        
        public string Raw_data_file_location { get => raw_data_file_location; set => raw_data_file_location = value; }

        public DataHandler(string path, int sheet)
        {
            ObservableCollection<RawRecord> records;
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
        }
        public int GetRowCount()
        {
            int row_count = 0;
            while (ws.Cells[(row_count+1), 1].Value2 != null)
            {
                row_count++;
            }
            return row_count;
        }
    }
}
