using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace HRS_ETL_Tool
{
    class AutoFitRowColumn
    {
        Application excel;
        Workbook wb;
        Worksheet ws;
        public void FittingRowColumn(string destinationpath)
        {
            try
            {
                excel = new Application();
                wb = excel.Workbooks.Open(destinationpath);
                excel.DisplayAlerts = false;
                ws = excel.ActiveSheet as Worksheet;
                ws.Columns.AutoFit();

                wb.Close(true, Type.Missing, Type.Missing);
            }
            catch
            {
                Program.CloseAllExcelAppication(ws, wb, excel);
            }
            finally
            {
                Program.CloseAllExcelAppication(ws, wb, excel);
            }
        }
    }
}
