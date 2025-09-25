using ExcelEf.Services;
using System;
using System.Data;
using System.Collections.Generic;
using ExcelEf.Models;
using ExcelEf.Mappers;
using System.Linq;
using System.Text;
using ExcelEf.Excel;

namespace ExcelEf
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {

            DateTime startDate = new DateTime(2025, 7, 1);
            DateTime endDate = new DateTime(2025, 8, 31);
            var before = DateTime.Now;

            //Выгрузка из бд
            var dbService = new DatabaseService();

            //Выгрузка параметров из бд
            var wellsTable = dbService.ExecuteQuery(SQLQueries.SQLGetWellID(startDate.ToString("dd.MM.yyyy"), endDate.ToString("dd.MM.yyyy")));
            List<Wells> wells = WellsDataMapper.ConvertToWellsList(wellsTable);
                        
             //Выгрузка ДНС
             var wellsDNS = dbService.ExecuteQuery(SQLQueries.SQLGetDNS());
             List<WellsDNS> dns = WellsDataMapper.ConvertToDNSList(wellsDNS);

             //Время после выгрузки ДНС и бд
            
                //Выгрузка ХАЛ
             var wellHal = dbService.ExecuteQuery(SQLQueries.SQLGetHAL(startDate.ToString("dd.MM.yyyy"), endDate.ToString("dd.MM.yyyy")));
             List<WellsHAL> hal = WellsDataMapper.ConvertToHalList(wellHal);
                //Выгрузка ХАЛ Тренд
            var wellTrend = dbService.ExecuteQuery(SQLQueries.SQLGetTrendHAL(startDate.ToString("dd.MM.yyyy"), endDate.ToString("dd.MM.yyyy")));
             List<WellsHAL> trend = WellsDataMapper.ConvertToHalList(wellTrend);

            var timeBeforeDbdns = DateTime.Now - before;
            Console.WriteLine($"Время после выгрузки БД : {timeBeforeDbdns}сек");

                GetExcelFormat.CreateExcelFile(wells, dns, hal, trend, startDate, endDate, "D:\\Отчёт.xlsx");
            Console.WriteLine("Готово!");

             var spendTime2 = DateTime.Now - before;
             Console.WriteLine($"Время выгрузки в Эксель: {spendTime2 - timeBeforeDbdns}сек");

                var spendTime = DateTime.Now - before;
             Console.WriteLine($"Время выполнения программы: {spendTime}сек");

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
                Console.ReadKey();
            }
    }
}
