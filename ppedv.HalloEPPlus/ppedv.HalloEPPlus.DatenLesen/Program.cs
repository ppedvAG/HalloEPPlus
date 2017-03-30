using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ppedv.HalloEPPlus.DatenLesen
{
    class Program
    {
        static void Main(string[] args)
        {
            //Exceldatei als FileInfo angeben
            string dateiname = "HalloExcel.xlsx";
            var fi = new FileInfo(dateiname);

            //dem ExcelPackage die FileInfo mit unserer Datei übergeben
            //zum lesen öffnet
            using (var pack = new ExcelPackage(fi))
            {
                //zugriff auf 
                var ws = pack.Workbook.Worksheets["Hallo"];

                foreach (var item in ws.Cells["B2:B40"].Where(x => x.GetValue<decimal>() > 50))
                {
                    Console.WriteLine(ws.Cells[item.Start.Row, item.Start.Column - 1].GetValue<DateTime>().ToShortDateString());
                }
            }

            Console.WriteLine("Ende");
            Console.ReadKey();
        }
    }
}
