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

                ws.Cells[1,1].Value = "Hallo A1";

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
