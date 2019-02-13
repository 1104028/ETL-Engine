using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Xml;
using System.IO;
using System.Collections.Generic;

namespace HRS_ETL_Tool
{
    class Source
    {
        public static int NumRows, SourceNumCols;
        public static string[,] sourcefile;
        public static string[,] headerformaterror = new string[2000, 3];
        public static int headerformaterror_rownumber;
        object[,] values;
        bool source_flag;
        bool separator_flg = false;
        static string[] Source_Field_Names = new string[] {
            "PROPCODE",
            "PROPNAME",
            "PROPADD1","PROPADD2",
            "PROPPOSTCODE",
            "PROPCITY",
            "PROPCOUNTRY",
            "MAINPHONECOUNTRY","MAINPHONECITY","MAINPHONE",
            "MAINFAXCOUNTRY","MAINFAXCITY","MAINFAX",
            "RFP_NAME",
            "RFP_EMAIL",
            "RFP_PHONECOUNTRYCODE","RFP_PHONECITYCODE","RFP_PHONE",
            "SEASON1START","SEASON1END",
                        "LRA_S1_RT1_SGL","LRA_S1_RT1_DBL","LRA_S1_RT2_SGL","LRA_S1_RT2_DBL","LRA_S1_RT3_SGL","LRA_S1_RT3_DBL",
                        "NLRA_S1_RT1_SGL","NLRA_S1_RT1_DBL","NLRA_S1_RT2_SGL","NLRA_S1_RT2_DBL","NLRA_S1_RT3_SGL","NLRA_S1_RT3_DBL",

            "SEASON2START","SEASON2END",
                        "LRA_S2_RT1_SGL","LRA_S2_RT1_DBL","LRA_S2_RT2_SGL","LRA_S2_RT2_DBL","LRA_S2_RT3_SGL","LRA_S2_RT3_DBL",
                        "NLRA_S2_RT1_SGL","NLRA_S2_RT1_DBL","NLRA_S2_RT2_SGL","NLRA_S2_RT2_DBL","NLRA_S2_RT3_SGL","NLRA_S2_RT3_DBL",

             "SEASON3START", "SEASON3END",
                        "LRA_S3_RT1_SGL","LRA_S3_RT1_DBL","LRA_S3_RT2_SGL","LRA_S3_RT2_DBL","LRA_S3_RT3_SGL","LRA_S3_RT3_DBL",
                        "NLRA_S3_RT1_SGL","NLRA_S3_RT1_DBL","NLRA_S3_RT2_SGL","NLRA_S3_RT2_DBL","NLRA_S3_RT3_SGL","NLRA_S3_RT3_DBL",

             "SEASON4START", "SEASON4END",
                        "LRA_S4_RT1_SGL","LRA_S4_RT1_DBL","LRA_S4_RT2_SGL","LRA_S4_RT2_DBL","LRA_S4_RT3_SGL","LRA_S4_RT3_DBL",
                        "NLRA_S4_RT1_SGL","NLRA_S4_RT1_DBL","NLRA_S4_RT2_SGL","NLRA_S4_RT2_DBL","NLRA_S4_RT3_SGL","NLRA_S4_RT3_DBL",

              "SEASON5START","SEASON5END",
                        "LRA_S5_RT1_SGL","LRA_S5_RT1_DBL","LRA_S5_RT2_SGL","LRA_S5_RT2_DBL","LRA_S5_RT3_SGL","LRA_S5_RT3_DBL",
                        "NLRA_S5_RT1_SGL","NLRA_S5_RT1_DBL","NLRA_S5_RT2_SGL","NLRA_S5_RT2_DBL","NLRA_S5_RT3_SGL","NLRA_S5_RT3_DBL",

              "BD1_START","BD1_END", "BD1_NAME",
                        "BD1_RT1_SGL","BD1_RT1_DBL","BD1_RT2_SGL","BD1_RT2_DBL","BD1_RT3_SGL","BD1_RT3_DBL",

              "BD2_START",  "BD2_END","BD2_NAME",
                        "BD2_RT1_SGL","BD2_RT1_DBL","BD2_RT2_SGL","BD2_RT2_DBL","BD2_RT3_SGL","BD2_RT3_DBL",

              "BD3_START","BD3_END", "BD3_NAME",
                        "BD3_RT1_SGL","BD3_RT1_DBL","BD3_RT2_SGL","BD3_RT2_DBL","BD3_RT3_SGL","BD3_RT3_DBL",

              "BD4_START",    "BD4_END","BD4_NAME",
                        "BD4_RT1_SGL","BD4_RT1_DBL","BD4_RT2_SGL","BD4_RT2_DBL","BD4_RT3_SGL","BD4_RT3_DBL",

              "BD5_START","BD5_END", "BD5_NAME",
                        "BD5_RT1_SGL","BD5_RT1_DBL","BD5_RT2_SGL","BD5_RT2_DBL","BD5_RT3_SGL","BD5_RT3_DBL",

              "BD6_START","BD6_END", "BD6_NAME",
                        "BD6_RT1_SGL","BD6_RT1_DBL","BD6_RT2_SGL","BD6_RT2_DBL","BD6_RT3_SGL","BD6_RT3_DBL",

              "BD7_START","BD7_END", "BD7_NAME",
                        "BD7_RT1_SGL","BD7_RT1_DBL","BD7_RT2_SGL","BD7_RT2_DBL","BD7_RT3_SGL","BD7_RT3_DBL",

              "BD8_START", "BD8_END","BD8_NAME",
                        "BD8_RT1_SGL","BD8_RT1_DBL","BD8_RT2_SGL","BD8_RT2_DBL","BD8_RT3_SGL","BD8_RT3_DBL",

              "BD9_START", "BD9_END","BD9_NAME",
                        "BD9_RT1_SGL","BD9_RT1_DBL","BD9_RT2_SGL","BD9_RT2_DBL","BD9_RT3_SGL","BD9_RT3_DBL",

              "BD10_START", "BD10_END","BD10_NAME",
                        "BD10_RT1_SGL","BD10_RT1_DBL","BD10_RT2_SGL","BD10_RT2_DBL","BD10_RT3_SGL","BD10_RT3_DBL",

            "BREAK_INCLUDE",
            "BREAK_FEE",
            "RATE_CURR",
            "LODGTX_INCLUDE",
            "STATETX_INCLUDE",
            "CITYTX_INCLUDE",
            "VATGSTRM_INCLUDE",
            "SERVICE_INCLUDE",
            "OCC_INCLUDE",
            "OTHERTX_FEE_INCL",
            "CANC_POL",
            "PARK_INCLUDE",
            "PARK_FEE",
            "WIRELESS_INCLUDE",
            "WIRELESS_FEE",
            "HSIA_INCLUDE",
            "HSIA_FEE",
            "AIRTRANS_INCLUDE",
            "AIRTRANS_FEE",
            "OFFTRANS_INCLUDE",
            "FITNESS_INCLUDE_ON",
            "FITNESS_FEE_ON",
            "LASTROOMAVAIL_BD"

        };

        static string[] Mandotory_Source_Fields = new string[]
        {
            "PROPNAME",
            "PROPADD1",
            "PROPADD2",
            "PROPCITY",
            "PROPCOUNTRY",
            "SEASON1START",
            "SEASON1END",
            "LRA_S1_RT1_SGL",
            "LRA_S1_RT1_DBL",
            "LRA_S1_RT2_SGL",
            "LRA_S1_RT2_DBL",
            "LRA_S1_RT3_SGL",
            "LRA_S1_RT3_DBL",
            "NLRA_S1_RT1_SGL",
            "NLRA_S1_RT1_DBL",
            "NLRA_S1_RT2_SGL",
            "NLRA_S1_RT2_DBL",
            "NLRA_S1_RT3_SGL",
            "NLRA_S1_RT3_DBL",
            "SEASON2START",
            "SEASON2END",
            "LRA_S2_RT1_SGL",
            "LRA_S2_RT1_DBL",
            "LRA_S2_RT2_SGL",
            "LRA_S2_RT2_DBL",
            "LRA_S2_RT3_SGL",
            "LRA_S2_RT3_DBL",
            "NLRA_S2_RT1_SGL",
            "NLRA_S2_RT1_DBL",
            "NLRA_S2_RT2_SGL",
            "NLRA_S2_RT2_DBL",
            "NLRA_S2_RT3_SGL",
            "NLRA_S2_RT3_DBL",
            "SEASON3START",
            "SEASON3END",
            "LRA_S3_RT1_SGL",
            "LRA_S3_RT1_DBL",
            "LRA_S3_RT2_SGL",
            "LRA_S3_RT2_DBL",
            "LRA_S3_RT3_SGL",
            "LRA_S3_RT3_DBL",
            "NLRA_S3_RT1_SGL",
            "NLRA_S3_RT1_DBL",
            "NLRA_S3_RT2_SGL",
            "NLRA_S3_RT2_DBL",
            "NLRA_S3_RT3_SGL",
            "NLRA_S3_RT3_DBL",
            "SEASON4START",
            "SEASON4END",
            "LRA_S4_RT1_SGL",
            "LRA_S4_RT1_DBL",
            "LRA_S4_RT2_SGL",
            "LRA_S4_RT2_DBL",
            "LRA_S4_RT3_SGL",
            "LRA_S4_RT3_DBL",
            "NLRA_S4_RT1_SGL",
            "NLRA_S4_RT1_DBL",
            "NLRA_S4_RT2_SGL",
            "NLRA_S4_RT2_DBL",
            "NLRA_S4_RT3_SGL",
            "NLRA_S4_RT3_DBL",
            "SEASON5START",
            "SEASON5END",
            "LRA_S5_RT1_SGL",
            "LRA_S5_RT1_DBL",
            "LRA_S5_RT2_SGL",
            "LRA_S5_RT2_DBL",
            "LRA_S5_RT3_SGL",
            "LRA_S5_RT3_DBL",
            "NLRA_S5_RT1_SGL",
            "NLRA_S5_RT1_DBL",
            "NLRA_S5_RT2_SGL",
            "NLRA_S5_RT2_DBL",
            "NLRA_S5_RT3_SGL",
            "NLRA_S5_RT3_DBL",
            "BD1_START",
            "BD1_END",
            "BD1_NAME",
            "BD1_RT1_SGL",
            "BD1_RT1_DBL",
            "BD1_RT2_SGL",
            "BD1_RT2_DBL",
            "BD1_RT3_SGL",
            "BD1_RT3_DBL",
            "BD2_START",
            "BD2_END",
            "BD2_NAME",
            "BD2_RT1_SGL",
            "BD2_RT1_DBL",
            "BD2_RT2_SGL",
            "BD2_RT2_DBL",
            "BD2_RT3_SGL",
            "BD2_RT3_DBL",
            "BD3_START",
            "BD3_END",
            "BD3_NAME",
            "BD3_RT1_SGL",
            "BD3_RT1_DBL",
            "BD3_RT2_SGL",
            "BD3_RT2_DBL",
            "BD3_RT3_SGL",
            "BD3_RT3_DBL",
            "BD4_START",
            "BD4_END",
            "BD4_NAME",
            "BD4_RT1_SGL",
            "BD4_RT1_DBL",
            "BD4_RT2_SGL",
            "BD4_RT2_DBL",
            "BD4_RT3_SGL",
            "BD4_RT3_DBL",
            "BD5_START",
            "BD5_END",
            "BD5_NAME",
            "BD5_RT1_SGL",
            "BD5_RT1_DBL",
            "BD5_RT2_SGL",
            "BD5_RT2_DBL",
            "BD5_RT3_SGL",
            "BD5_RT3_DBL",
            "BD6_START",
            "BD6_END",
            "BD6_NAME",
            "BD6_RT1_SGL",
            "BD6_RT1_DBL",
            "BD6_RT2_SGL",
            "BD6_RT2_DBL",
            "BD6_RT3_SGL",
            "BD6_RT3_DBL",
             "BD7_START",
            "BD7_END",
            "BD7_NAME",
             "BD7_RT1_SGL",
            "BD7_RT1_DBL",
            "BD7_RT2_SGL",
            "BD7_RT2_DBL",
            "BD7_RT3_SGL",
            "BD7_RT3_DBL",
            "BD8_START",
            "BD8_END",
            "BD8_NAME",
             "BD8_RT1_SGL",
            "BD8_RT1_DBL",
            "BD8_RT2_SGL",
            "BD8_RT2_DBL",
            "BD8_RT3_SGL",
            "BD8_RT3_DBL",
            "BD9_START",
            "BD9_END",
            "BD9_NAME",
            "BD9_RT1_SGL",
            "BD9_RT1_DBL",
            "BD9_RT2_SGL",
            "BD9_RT2_DBL",
            "BD9_RT3_SGL",
            "BD9_RT3_DBL",
            "BD10_START",
            "BD10_END",
            "BD10_NAME",
            "BD10_RT1_SGL",
            "BD10_RT1_DBL",
            "BD10_RT2_SGL",
            "BD10_RT2_DBL",
            "BD10_RT3_SGL",
            "BD10_RT3_DBL",
            "BREAK_INCLUDE",
            "RATE_CURR"
        };

        static int[] Mandtory_Map = new int[167];
        public static int[] Map = new int[201];
        static Application source_excel = new Application();

        static Workbook source_wb;
        static Worksheet source_ws;
        public static int source_total_rows;

        public void OpenSourceExcel()
        {
            source_flag = true;

            string fileName = System.IO.Path.GetFileName(Program.sourcepath);
            var extension = Path.GetExtension(fileName);
            var fname = Path.GetFileNameWithoutExtension(fileName);

            try
            {
                source_excel.DisplayAlerts = false;
                source_wb = source_excel.Workbooks.Open(Program.sourcepath);
                source_ws = source_wb.Worksheets[1];

                for (int i = 0; i < 201; i++)
                {
                    Map[i] = -10;
                }
                for (int i = 0; i < 167; i++)
                {
                    Mandtory_Map[i] = -1;
                }

                ReadSourceExcel();

                source_wb.Close(0);
            }
            catch (Exception e)
            {
                source_flag = false;
                Console.WriteLine(e.Message);

                Program.CloseAllExcelAppication(source_ws, source_wb, source_excel);
            }
            finally
            {
                Program.CloseAllExcelAppication(source_ws, source_wb, source_excel);
            }

            if (source_flag)
            {
                SaveSourceData();
            }
            else
            {
                LogFile log = new LogFile();
                log.FormatErrorLog();
                Program.continuexecution = false;
            }
        }
        public void ReadSourceExcel()
        {
            Range range = source_ws.UsedRange;
            values = (object[,])range.Value;
            SourceNumCols = values.GetLength(1);

            for (int i = 0; i < 167; i++)
            {
                for (int j = 0; j < 3; j++)
                    headerformaterror[i, j] = "";
            }

            //find header column number in source excel;
            for (int i = 0; i < Source_Field_Names.Length; i++)
            {
                for (int j = 1; j <= SourceNumCols; j++)
                {
                    if (values[3, j] != null)
                    {
                        if (Source_Field_Names[i].Equals(values[3, j].ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            Map[i] = j;
                            break;
                        }
                    }

                }
                if (Map[i] == -10) Map[i] = -1;
            }

            //find mandatory fields column number in source excel;
            for (int i = 0; i < Mandotory_Source_Fields.Length; i++)
            {
                for (int j = 1; j <= SourceNumCols; j++)
                {
                    if (values[3, j] != null)
                    {
                        if (Mandotory_Source_Fields[i].Equals(values[3, j].ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            Mandtory_Map[i] = j;
                            break;
                        }
                    }
                }
            }

            //headerformaterror format error log file header
            headerformaterror[0, 0] = "Field Name";
            headerformaterror[0, 1] = "Field Type";
            headerformaterror[0, 2] = "Status";

            headerformaterror_rownumber = 1;

            for (int i = 0; i < 167; i++)
            {
                if (Mandtory_Map[i] == -1)
                {
                    if (!separator_flg)
                    {
                        headerformaterror[headerformaterror_rownumber, 0] = Mandotory_Source_Fields[i];
                        headerformaterror[headerformaterror_rownumber, 1] = "Mandatory";
                        headerformaterror[headerformaterror_rownumber, 2] = "Missing";
                        separator_flg = true;
                    }
                    else
                    {
                        headerformaterror[headerformaterror_rownumber, 0] = Mandotory_Source_Fields[i];
                        headerformaterror[headerformaterror_rownumber, 1] = "Mandatory";
                        headerformaterror[headerformaterror_rownumber, 2] = "Missing";
                    }
                    headerformaterror_rownumber++;
                    source_flag = false;
                }
            }
        }
        public void SaveSourceData()
        {
            NumRows = values.GetLength(0);
            SourceNumCols = values.GetLength(1);
            sourcefile = new string[NumRows + 1, SourceNumCols + 1];

            source_total_rows = 4;

            for (int i = 4; i <= NumRows; i++)
            {
                if (Is_Not_EmptyRows(i) == true)
                {
                    for (int j = 1; j <= SourceNumCols; j++)
                    {
                        if (values[i, j] != null)
                        {
                            sourcefile[source_total_rows, j] = values[i, j].ToString();
                        }
                        else sourcefile[source_total_rows, j] = "";
                    }
                    source_total_rows++;
                }
            }
            source_total_rows = source_total_rows - 1;
        }
        public bool Is_Not_EmptyRows(int row_num)
        {
            var col_num = 0;

            for (int j = 1; j <= SourceNumCols; j++)
            {

                if (values[row_num, j] == null)
                {
                    col_num++;
                }
                else
                    break;
            }
            if (col_num == SourceNumCols)
                return false;
            else
                return true;
        }
    }
}
