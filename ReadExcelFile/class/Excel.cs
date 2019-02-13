using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelFile
{
    public class Excel
    {
        public string Date;
        public string Time;
        public string HomeCharA;
        public string HomeCharV;
        public string SolarCharA;
        public string SolarCharV;
        public string UseA;
        public string UseV;
        public string Distanc;
        public string Power;
        public string Energy;
    }

    public class CalculatePower
    {
        public string Time;
        public double? Power; 
    }

    public class ListCalPower
    {
        public List<CalculatePower> calObj;
    }
}
