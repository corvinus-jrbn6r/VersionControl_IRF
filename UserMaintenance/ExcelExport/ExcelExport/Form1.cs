using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace ExcelExport
{
    public partial class Form1 : Form
    {
        RealEstateEntities context = new RealEstateEntities();
        List<Flat> lakasok;

        Excel.Application xlApp; // A Microsoft Excel alkalmazás
        Excel.Workbook xlWB; // A létrehozott munkafüzet
        Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül
        string[] headers;


        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();


        }

        public void LoadData()
        {
            lakasok = context.Flats.ToList();
        }

        public void CreateExcel()
        {
            try
            {

                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();

                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value); // mert egy üres doksit hozunk létre

                // Új munkalap, az a munkalap, ami aktuálisan meg van nyitva (amikor új excelt nyitok, lesz egy új munkalapunk)
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                CreateTable(); // Ennek megírása a következő feladatrészben következik
                FormatTable();

                // Control átadása a felhasználónak - miután megnyitottuk és írtunk a munkalapra
                xlApp.Visible = true; //megjelenítjük az eddig háttérben futó alkalmazást
                xlApp.UserControl = true; //átadjuk a vezérlőt
            }
            catch (Exception ex) // Hibakezelés a beépített hibaüzenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error"); //\nLine a sortörés

                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing); //bezárom a munkafüzetet
                //első: mentsük el a hibás állomány, hova mentsem, ez is hogy hova mentsem)
                xlApp.Quit(); //ezzel lehet teljesen lezárni a dolgot, bezárom az appot
                xlWB = null; //törlöm a munkafüzetet
                xlApp = null; //törlöm az appot
            }
            /*finally //akkor is lefut, ha a catch-el elkaptunk egy hibát
            {
                MessageBox jöhet ide
            }*/


        }
        private void CreateTable()
        {
            // ha nem sorolom fel, akkor meg kell adni a számot a kapcsos zárójelben
           headers = new string[] {
             "Kód",
             "Eladó",
             "Oldal",
             "Kerület",
             "Lift",
             "Szobák száma",
             "Alapterület (m2)",
             "Ár (mFt)",
             "Négyzetméter ár (Ft/m2)"};

            for (int i = 0; i < headers.Length; i++)
                xlSheet.Cells[1, i + 1] = headers[i];

            // xlSheet.Cells[1, 1] = headers[0];  ezzel lehetne beírni egy cellába, a fentebbivel az össszesbe

            object[,] values = new object[lakasok.Count, headers.Length]; //így mégcsak VS oldalon van

            int counter = 0;
            int floorcolumn = 6;
            foreach (var lakas in lakasok)
            {
                values[counter, 0] = lakas.Code; // 0-tól indexálom az oszlopokat
                values[counter, 1] = lakas.Vendor;
                values[counter, 2] = lakas.Side;
                values[counter, 3] = lakas.District;

                if (lakas.Elevator)
                    values[counter, 4] = "Van";
                else
                    values[counter, 4] = "Nincs";

                values[counter, 5] = lakas.NumberOfRooms;
                values[counter, 6] = lakas.FloorArea;
                values[counter, 7] = lakas.Price;
                values[counter, 8] = string.Format("={0}/{1}*{2}", "H" + (counter + 2).ToString(),
                    GetCell((counter + 2), floorcolumn + 1), 1000000);
                //"=H2/G2*1000000";

                counter++;
            }
            //ezzel egy range-et választottam ki a sheeten
            var range = xlSheet.get_Range(
                GetCell(2, 1),
                GetCell(1 + values.GetLength(0), values.GetLength(1)));

            //value 2 kell az excelhez, amikor object valami
            range.Value2 = values;
        }

        private string GetCell(int x, int y) //ezt nem kell tudnom, de fel kell tudnom használni (tudnom kell, mi ez)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }

        private void FormatTable(){
            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range completeTableRange = xlSheet.get_Range(GetCell(1, 1), GetCell(lastRowID, headers.Length));
            completeTableRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range firstRowRange = xlSheet.get_Range(GetCell(1, 1), GetCell(lastRowID, 1));
            firstRowRange.Font.Bold = true;
            firstRowRange.Interior.Color = Color.LightYellow;

            Excel.Range lastRowRange = xlSheet.get_Range(GetCell(1,headers.Length), GetCell(lastRowID, headers.Length));
            lastRowRange.Interior.Color = Color.LightGreen;
            lastRowRange.NumberFormat = "###,###.00";

        }
    }
}
