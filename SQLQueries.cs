using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEf.Services
{
    internal class SQLQueries
    {

        public static string SQLGetWellID(string StartDate, string EndDate)
        {
            return "select v.measure_date, t.*,v.measure_type_id, v.value, tex.liq_rate, tex.oil_rate, tex.water_cut " +
              "from (SELECT a.well_id, " +
                  "w.field_name, " +
                  "w.field_code, " +
                  "w.well_name, " +
                  "w.well_cluster, " +
                  "w.gsu_name, " +
                  "w.shop_name, " +
                  "w.team_no, " +
                         "(select s.ns_1 " +
                         "from oilinfo.class$ s " +
                         "where s.cd_1 = a.prod_method_id) tek_sposob, " +
                         "'В работе' sost " +
                         "FROM v_well_list " +
                         "a, " +
                         "measure_journal_well b, " +
                         "fund_group_status    c, " +
                         "v_well_full          w " +
                         "WHERE a.well_id = b.well_id " +
                         "AND b.journal_id = 176956 " +
                         "AND a.status_id = c.status_id(+) " +
                         "AND a.well_id = w.well_id(+) " +
                         "and c.status_id = 'SS0001' " +
                         "and a.purpose_id = 'XR0011' " +
                         "order by a.well_id) t " +
                         "left join (select tr.well_id,tr.liq_rate,tr.water_cut, tr.oil_rate " +
                         "from MV_WELL_OP tr " +
                         "where tr.calc_date = trunc(add_months(sysdate-1, -1), 'MM')) tex " +
                         "on t.well_id = tex.well_id " +
                             "left join (select m.well_id, m.measure_date, m.value, m.measure_type_id " +
                             "from well_measure m " +
                             "where m.measure_type_id in (1,3,4,12) " + //--QG = 1,  Обв = 3, Qn = 4, Обв Хад = 11, Qg ТМ = 12,  Тренд обв ХАЛ = 213
                             "and m.measure_date >=  '" + StartDate + "' " +
                             "and m.measure_date <= '" + EndDate + "') v " +
                             "on t.well_id = v.well_id " +
                 "order by v.well_id, v.measure_date ";
        }
        public static string SQLGetDNS()
        {
            return "select h.well_id, h.dns_name from fo_hsod h " + 
                "where h.dt = trunc(sysdate)";
        }
        public static string SQLGetHAL(string StartDate, string EndDate)
        {
            return "select " +
                "wm.well_id, " +
                "trunc(metering_date) measure_date, " +
                "Round(AVG(value),1) as VALUE " +
                "from" +
                " V_HAL_WATER_RESEARCH_CONFIRMED wm" +
                " where" +
                " metering_date >= '" + StartDate + "'" +
                " and metering_date < '" + EndDate + "'" +
                " AND 1=1 " +
                "group by metering_date, well_id " +
                "order by wm.well_id, wm.metering_date"; 
        }
        public static string SQLGetTrendHAL(string StartDate, string EndDate)
        {
            return "select " +
                "wm.well_id, " +
                "trunc(metering_date) measure_date, " +
                "Round(AVG(trend),1) as VALUE " +
                "from" +
                " v_hal_water_research wm" +
                " where" +
                " metering_date >= '" + StartDate + "'" +
                " and metering_date < '" + EndDate + "'" +
                " AND 1=1 " +
                "group by metering_date, well_id " +
                "order by wm.well_id, wm.metering_date";
        }
    }
}
