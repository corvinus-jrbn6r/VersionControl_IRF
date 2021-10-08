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

        }

    }
}
