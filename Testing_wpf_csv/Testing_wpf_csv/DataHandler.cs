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

        _Application excel;
        Workbook wb;
        Worksheet ws;
        


        public string Raw_data_file_location { get => raw_data_file_location; set => raw_data_file_location = value; }

        //init instance, and open a new excel app
        public DataHandler()
        {
            excel = new _Excel.Application();//start excel app            
        }
        //read cell content, returns a double
        public double ReadCell(int cells_down, int cells_to_right)
        {
            if (ws.Cells[cells_down, cells_to_right].Value2 != null)
            {
                return ws.Cells[cells_down, cells_to_right].Value2;
            }
            else return -1.0;
        }
        public string ReadCell_str(int cells_down, int cells_to_right)
        {
            if (ws.Cells[cells_down, cells_to_right].Value2 != null)
            {
                return ws.Cells[cells_down, cells_to_right].Value2;
            }
            else return "";
        }
        //write to excel cell 
        public void WriteToCell(int cells_down, int cells_to_right, string s)
        {
            ws.Cells[cells_down, cells_to_right].Value2 = s;
        }
        public void WriteFormulaToCell(int cells_down, int cells_to_right, string formula)
        {
            
            ws.Cells[cells_down, cells_to_right].Formula = formula;            
        }
        
        public void WriteToCell(int cells_down, int cells_to_right, int num)
        {
            ws.Cells[cells_down, cells_to_right].Value2 = num;
        }
        public void WriteToCell(int cells_down, int cells_to_right, double num)
        {
            ws.Cells[cells_down, cells_to_right].Value2 = num;
        }
        //calculates how many rows are used, count stops when an empty cell is found  
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

        //Create new excel workbook
        public void CreateNewFile(string path)
        {
            wb = excel.Workbooks.Add();            
            wb.SaveAs(path);            
        }
        //will style and draw graph
        public void StyleExcelFile(int record_count)
        {
            object misValue = System.Reflection.Missing.Value;
            _Excel.Range bold,percentage,subtitle,chart_range;
            string end_cell,x_end_cell;//used to get Range object with data for graph
            Worksheet temp_ws = (_Excel.Worksheet)wb.Worksheets.get_Item(2);//temp worksheet to get range with data for graph
            int end_row = record_count + 1;
            //int end_row = 200;//for tests only
            end_cell = "b" + end_row.ToString();
            x_end_cell = "a" + end_row.ToString();
            chart_range = temp_ws.Range["a2", end_cell];
            _Excel.ChartObjects xlCharts = (_Excel.ChartObjects)ws.ChartObjects(Type.Missing);

            _Excel.ChartObject myChart = (_Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);            
            _Excel.Chart chartPage = myChart.Chart;
            
            chartPage.SetSourceData(chart_range, misValue);
            chartPage.ChartType = _Excel.XlChartType.xlLine;            
            bold = ws.Range["b7", "j7"];
            subtitle = ws.Range["b4"];
            chartPage.HasTitle = true;
            chartPage.ChartTitle.Text = "Average Canopy Height/Footage Time (meters/seconds)";
                        

            
            
            bold.Font.Bold = true;
            percentage = ws.Range["c9","g9"];
            percentage.NumberFormat = "#.00%";            
            percentage.Font.Size = 14;
            subtitle.Font.Size = 14;            
        }

        //wb must be initialized already
        public void CreateNewSheet()
        { 
            wb.Worksheets.Add(After: ws);            
        }
        //Selects a new sheet, further changes and actions will be done to selected sheet
        public void SelectSheet(int sheet_index)
        {
            ws = (_Excel.Worksheet)wb.Worksheets.get_Item(sheet_index);
        }

        public void SelectSheet(string sheet_index)
        {
            ws = (_Excel.Worksheet)wb.Worksheets.get_Item(sheet_index);
        }
        //this sheet will be displayed when user opens excel document
        public void SelectOpenSheet()
        {
            ws.Select();
        }

        //rename a already initialized excel worksheet
        public void RenameSheet(string sheet_name)
        {
            ws.Name = sheet_name;
        }
        //
        public void Open(string path, int sheet)
        {            
            wb = excel.Workbooks.Open(path);
            SelectSheet(sheet);
        }
        public void Open(string path, string sheet)
        {
            wb = excel.Workbooks.Open(path);
            SelectSheet(sheet);
        }
        public List<string> GetRawDataList(string csv_path)
        {
            return File.ReadAllLines(csv_path).ToList();
        }
        
        //Save changes to excel workbook file,wb must already be initialized
        public void SaveWorkBookProgress()
        {
            wb.Save();
        }
        //
           
        public void Close()
        {
            wb.Close();            
        }

        
    }

}
