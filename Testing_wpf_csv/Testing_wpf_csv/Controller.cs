﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Testing_wpf_csv.Control
{
    class Controller
    {
        //private attributes
        private List<FileController> files;

        //properties
        public List<FileController> Files { get => files; set => files = value; }
        
        public Controller(string[] file_path)
        {

            files = new List<FileController>();//init list, with no items          
            
            foreach (string path in file_path)
            {
                files.Add(new FileController(path));//add a file controller for every raw data csv found to list
            }
        }

        /* 
         *  
        */
        public void ProcessFiles()
        {
            int count = 0;
            foreach (var file_controller in files)
            {
                if (file_controller.IsValid)
                {
                    //(sender as BackgroundWorker).ReportProgress(count, file_controller.Exported_file_path);
                    file_controller.ProcessFile();
                    count++;
                }
            }
        }
    }
}
