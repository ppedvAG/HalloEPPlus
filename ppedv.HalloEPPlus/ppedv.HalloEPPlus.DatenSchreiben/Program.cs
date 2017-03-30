using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ppedv.HalloEPPlus.DatenSchreiben
{
    class Program
    {
        static void Main(string[] args)
        {
            //Exceldatei als FileInfo angeben
            string dateiname = "HalloExcel.xlsx";
            var fi = new FileInfo(dateiname);

            //wenn Datei existiert, dann löschen
            if (fi.Exists)
                fi.Delete();

            //dem ExcelPackage die FileInfo mit unserer Datei übergeben
            //wird bei speichern neu angelegt oder falls existiert zum lesen geöffnet
            using (var pack = new ExcelPackage(fi))
            {
                //neues WorkSheet erstellen mit den Titel "Hallo"
                var ws = pack.Workbook.Worksheets.Add("Hallo");

                //ws.Cells["A1"].Value = "Hallo A1";
                //ws.Cells[1,1].Value = "Hallo A1";

                ws.Cells[1, 1].Value = "Tag";
                ws.Cells[1, 2].Value = "Umsatz";

                var now = DateTime.Now;
                var ran = new Random();
                for (int i = 1; i <= DateTime.DaysInMonth(now.Year, now.Month); i++)
                {
                    DateTime day = new DateTime(now.Year, now.Month, i);
                    int rowNum = i + 1;

                    if (day.DayOfWeek == DayOfWeek.Saturday)
                    {
                        ws.Cells[rowNum, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[rowNum, 1].Style.Fill.BackgroundColor.Indexed = 41;
                    }

                    if (day.DayOfWeek == DayOfWeek.Sunday)
                    {
                        ws.Cells[rowNum, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[rowNum, 1].Style.Fill.BackgroundColor.Indexed = 42;
                    }


                    ws.Cells[rowNum, 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    ws.Cells[rowNum, 1].Value = day;

                    ws.Cells[rowNum, 2].Style.Numberformat.Format = "€#,##0.00";
                    ws.Cells[rowNum, 2].Value = ran.NextDouble() * 100;

                }

                ws.Column(1).AutoFit();
                ws.Column(2).AutoFit();


                //Änderungen der Datei abspeichern
                pack.Save();
            }

            //Datei mit dem Standardprogramm für .xlsx Dateien starten
            Process.Start(dateiname);


            Console.WriteLine("Ende");
            Console.ReadKey();
        }
    }
}
