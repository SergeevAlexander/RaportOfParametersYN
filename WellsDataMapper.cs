using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Threading.Tasks;
using ExcelEf.Models;

namespace ExcelEf.Mappers
{
    internal class WellsDataMapper
    {   //Запрос для формирование коллекции с скважинами
        public static List<Wells> ConvertToWellsList(List<Dictionary<string, object>> data)
        {
            var wells = new List<Wells>();
            foreach (var row in data)
            {
                wells.Add(new Wells
                {
                    WellId = GetIntValue(row, "WELL_ID"),
                    FieldName = GetStringValue(row, "FIELD_NAME"),
                    FieldCode = GetIntValue(row, "FIELD_CODE"),
                    WellName = GetStringValue(row, "WELL_NAME"),
                    WellCluster = GetStringValue(row, "WELL_CLUSTER"),
                    GsuName = GetStringValue(row, "GSU_NAME"),
                    ShopName = GetStringValue(row, "SHOP_NAME"),
                    TeamNo = GetStringValue(row, "TEAM_NO"),
                    TRQg = GetDoubleValue(row, "LIQ_RATE"),
                    TRQn = GetDoubleValue(row, "OIL_RATE"),
                    TRObv = GetDoubleValue(row, "WATER_CUT"),
                    TekSposob = GetStringValue(row, "TEK_SPOSOB"),
                    Sost = GetStringValue(row, "SOST"),
                    MeasureTypeId = GetIntValue(row, "MEASURE_TYPE_ID"),
                    Value = GetDoubleValue(row, "VALUE"),
                    MeasureDate = GetDateTimeValue(row, "MEASURE_DATE")
                });
            }
            return wells;
        }
        public static List<WellsDNS> ConvertToDNSList(List<Dictionary<string, object>> data)
        {
            var wells = new List<WellsDNS>();
            foreach (var row in data)
            {
                wells.Add(new WellsDNS
                {
                    WellId = GetIntValue(row, "WELL_ID"),
                    DNSName = GetStringValue(row, "DNS_NAME"),
                });
            }
            return wells;
        }
        public static List<WellsHAL> ConvertToHalList(List<Dictionary<string, object>> data)
        {
            var wells = new List<WellsHAL>();
            foreach (var row in data)
            {
                wells.Add(new WellsHAL
                {
                    WellId = GetIntValue(row, "WELL_ID"),
                    MeasureDate = GetDateTimeValue(row, "MEASURE_DATE"),
                    Value = GetDoubleValue(row, "VALUE"),

                });
            }
            return wells;
        }
        private static int GetIntValue(Dictionary<string,object> row, string key)
        {
            if (row.TryGetValue(key, out object value) && value != null)
            {
                if (value is int intValue) return intValue;
                if (int.TryParse(value.ToString(), out int parsed)) return parsed;
            }
            return 0;
        }
        private static string GetStringValue(Dictionary<string, object> row, string key)
        {
            if (row.TryGetValue(key, out object value) && value != null)
            {
                return value.ToString().Trim();
            }
            return String.Empty;
        }
        private static double GetDoubleValue(Dictionary<string, object> row, string key)
        {
            if (row.TryGetValue(key, out object value) && value != null)
            {
                if (value is double doubleValue) return doubleValue;
                if (double.TryParse(value.ToString(), out double parsed)) return parsed;
            }
            return 0;
        }
        private static DateTime? GetDateTimeValue(Dictionary<string, object> row, string key)
        {
            if (row.TryGetValue(key, out object value) && value != null)
            {
                if (value is DateTime dateValue) return dateValue.Date;
                if (DateTime.TryParse(value.ToString(), out DateTime parsed)) return parsed.Date;
            }
            return null;
        }

    }

}
