using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace HRS_ETL_Tool
{
    class OutputFileGenerate
    {
        Application merged_excel;
        Workbook merged_wb;
        Worksheet merged_ws;
        Range merged_xlRange;
        string destinationpath;
        string[,] output = new string[2000, 63];

        public void WriteOutputExcel()
        {
            output[0, 0] = Format.headervalue[0, 0];

            for (int i = 0; i < 63; i++)
            {
                output[1, i] = Format.headervalue[1, i];
            }
            int k = 2;

            for (int i = 0; i < DataMerge.destrow; i++)
            {
                for (int j = 0; j < 63; j++)
                {
                    output[k, j] = Format.destinationfile[i, j];
                }
                k++;

            }

            string fileName = System.IO.Path.GetFileName(Program.sourcepath);
            var extension = Path.GetExtension(fileName);
            var fname = Path.GetFileNameWithoutExtension(fileName);

            if (!System.IO.Directory.Exists(Program.destinationfolder))
            {
                System.IO.Directory.CreateDirectory(Program.destinationfolder);
            }

            destinationpath = Program.destinationfolder + "Output_ " + Program.fileno + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

            try
            {
                merged_excel = new Application();
                merged_wb = merged_excel.Workbooks.Add(1);
                merged_excel.DisplayAlerts = false;
                merged_ws = merged_wb.Worksheets[1];

                merged_xlRange = merged_ws.UsedRange;
                int merged_cols = merged_xlRange.Columns.Count;
                merged_cols = Format.formatNumCols;

                var startCell = (Range)merged_ws.Cells[1, 1];

                var endCell = (Range)merged_ws.Cells[DataMerge.destrow + 2, merged_cols];
                var writeRange = merged_ws.Range[startCell, endCell];

                var columnHeadingsRange = merged_ws.Range[merged_ws.Cells[1, 1], merged_ws.Cells[1, merged_cols]];
                columnHeadingsRange.Interior.Color = 0x9C632A;
                columnHeadingsRange.Font.Color = XlRgbColor.rgbWhite;
                columnHeadingsRange.Font.Size = 13;

                var columnHeadingsRangeHeader = merged_ws.Range[merged_ws.Cells[2, 1], merged_ws.Cells[2, merged_cols]];
                columnHeadingsRangeHeader.Interior.Color = 0xD9D9D9;
                columnHeadingsRangeHeader.Font.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeHeader.Font.Size = 12;
                columnHeadingsRangeHeader.EntireRow.Font.Bold = true;
                columnHeadingsRangeHeader.Borders.Color = XlRgbColor.rgbBlack;
                columnHeadingsRange.EntireColumn.AutoFit();
                columnHeadingsRangeHeader.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

                var columnHeadingsRangeData = merged_ws.Range[merged_ws.Cells[3, 1], merged_ws.Cells[DataMerge.destrow + 2, merged_cols]];
                columnHeadingsRangeData.Font.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeData.Font.Size = 11;
                columnHeadingsRangeData.Borders.Color = XlRgbColor.rgbBlack;

                var columnHeadingsRangeEmail = merged_ws.Range[merged_ws.Cells[3, 11], merged_ws.Cells[DataMerge.destrow + 2, 11]];
                columnHeadingsRangeEmail.Font.Color = XlRgbColor.rgbBlue;

                writeRange.Value = output;

                merged_wb.SaveAs(destinationpath);
                merged_wb.Close();
            }
            catch (Exception e)
            {
                Program.continuexecution = false;
                Console.WriteLine("Can't create target file because of these reasons:");
                Console.WriteLine(e.Message);
                Program.CloseAllExcelAppication(merged_ws, merged_wb, merged_excel);
            }
            finally
            {
                Program.CloseAllExcelAppication(merged_ws, merged_wb, merged_excel);
            }

            // AutoFit Columns
            AutoFitRowColumn fit = new AutoFitRowColumn();
            fit.FittingRowColumn(destinationpath);

            Program.fileno++;
        }
    }
}
