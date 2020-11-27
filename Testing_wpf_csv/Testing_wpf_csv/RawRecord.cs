using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Testing_wpf_csv.Models
{
    class RawRecord
    {
        private double time;
        private double average_shoot_height;
        private double latitude;
        private double longitud;
        private double smoothered_height;
        private double average_average_height;

        public double Time { get => time; set => time = value; }
        public double Average_shoot_height { get => average_shoot_height; set => average_shoot_height = value; }
        public double Latitude { get => latitude; set => latitude = value; }
        public double Longitud { get => longitud; set => longitud = value; }
        public double Smoothered_height { get => smoothered_height; set => smoothered_height = value; }
        public double Average_average_height { get => average_average_height; set => average_average_height = value; }
        public RawRecord(double time, double average_shoot_height, double latitud, double longitud )
        {
            this.time = time;
            this.average_shoot_height = average_shoot_height;
            this.latitude = latitud;
            this.longitud = longitud;
        }
    }
}
