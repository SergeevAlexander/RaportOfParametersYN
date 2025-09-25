using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEf.Models
{
    public class Wells
    {
        public int WellId { get; set; }
        public string FieldName { get; set; }
        public int FieldCode { get; set; }
        public string WellName { get; set; }
        public string WellCluster { get; set; }
        public string GsuName { get; set; }
        public string ShopName { get; set; }
        public string TeamNo { get; set; }
        public string TR { get; set; }
        public double TRQg { get; set; }
        public double TRQn { get; set; }
        public double TRObv { get; set; }
        public string TekSposob { get; set; }
        public string Sost { get; set; }
        public int MeasureTypeId { get; set; }
        public double Value { get; set; }
        public DateTime? MeasureDate { get; set; }

    }
    public class WellsDNS
    {
        public int WellId { get; set; }
        public string DNSName { get; set; }
    }

    public class WellsHAL
    {
        public int WellId { get; set; }
        public DateTime? MeasureDate { get; set; }
        public double Value { get; set; }
    }
}
