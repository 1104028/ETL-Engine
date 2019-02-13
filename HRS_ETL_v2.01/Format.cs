using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace HRS_ETL_Tool
{
    class Format
    {
        object[,] headerFormatValue;
        public static int formatNumCols;
        Worksheet fomat_ws;
        public static string[,] destinationfile = new string[2000, 64];
        Application fomat_excel;
        Workbook fomat_wb;
        public static bool format_flag;
        public static string[,] headervalue = new string[2, 64];

        public void OpenFormatExcel()
        {
            format_flag = true;// initially format flag true

            fomat_excel = new Application();// create an excel application

            try
            {
                fomat_excel.DisplayAlerts = false;// hide any alerts
                fomat_wb = fomat_excel.Workbooks.Open(Program.fomatpath);// open format excel files
                fomat_ws = fomat_wb.Worksheets[1];// open worksheet

                ReadFormatExcel();//call for read format excel file

                fomat_wb.Close(0);// close format workbook
            }
            catch
            {
                format_flag = false;// when openning format excel file if any error occurs set format flag false 
                Console.WriteLine("Format File can't Open, Please Check This File is Valid.");

                Program.CloseAllExcelAppication(fomat_ws, fomat_wb, fomat_excel);// close format excel file
            }
            finally
            {
                Program.CloseAllExcelAppication(fomat_ws, fomat_wb, fomat_excel);// close format excel file
            }

            if (format_flag)// if format flag is true
            {
                SaveFormatData();// call for save format excel data
            }
            else
                Program.UnsuccessfullyFinishExecution(); //if format flag is false then stop processing source files
        }
        // read format excel
        public void ReadFormatExcel()
        {
            try
            {
                Range rangev = fomat_ws.UsedRange;// create a range for format excel worksheet
                headerFormatValue = (object[,])rangev.Value2;// assign in memory or variable
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        //save format excel data
        public void SaveFormatData()
        {
            try
            {
                formatNumCols = headerFormatValue.GetLength(1);// get format excel column number 
                headervalue[0, 0] = headerFormatValue[1, 1].ToString();// assign output excel file header row 1

                // assign output excel file header row 2
                for (int j = 1; j <= formatNumCols; j++)
                {
                    try
                    {
                        headervalue[1, j - 1] = headerFormatValue[2, j].ToString();// save data header column names
                    }
                    catch
                    {
                        Console.WriteLine("In Format File, Header Field Can't be Null." + "\nThe Row and Column is: " + "2 " + j);
                        Program.UnsuccessfullyFinishExecution();//call for stop processing source files
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
