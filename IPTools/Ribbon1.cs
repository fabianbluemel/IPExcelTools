using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;


namespace IPTools
{
    
    public partial class Ribbon1
    {
        public string DNSServer01;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //< open Excel Worksheet >
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //< open Excel Worksheet >

            //< get Cell.Value >
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;

            actCell.Value2 = "=GetDNS(\"82.165.229.83\")";

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //< open Excel Worksheet >
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //< open Excel Worksheet >

            //< get Cell.Value >
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;

            actCell.Value2 = "=GetIP(\"web.de\")";
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //< open Excel Worksheet >
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //< open Excel Worksheet >

            //< get Cell.Value >
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;

            actCell.Value2 = "=PING(\"8.8.8.8\")";
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //< open Excel Worksheet >
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //< open Excel Worksheet >

            //< get Cell.Value >
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;

            actCell.Value2 = "=GetIPWSERVER()";
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            //< open Excel Worksheet >
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //< open Excel Worksheet >

            //< get Cell.Value >
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;

            actCell.Value2 = "=GetDNSWSERVER()";
        }
    }
}
