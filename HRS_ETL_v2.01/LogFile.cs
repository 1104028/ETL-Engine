using System;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace HRS_ETL_Tool
{
    class LogFile
    {
        public Application log_excel, error_log_excel, header_error_log_excel;
        public Workbook log_wb, error_log_wb, header_error_log_wb;
        public Worksheet log_ws, error_log_ws, header_error_log_ws;

        public object[,] values;
        public static string[,] logfilearray = new string[10000, 4];
        string fname, destFile, extension, savefilenameas;

        public void SuccessfulLog()
        {
            int notsuccessful = Source.source_total_rows - Program.successful_row - 3;
            if (notsuccessful== 0 && Program.lengthtoolong == true)// 
            {
                log_excel = new Application();
                try
                {
                    log_excel.DisplayAlerts = false;

                    if (!System.IO.File.Exists(Program.logfile))
                    {
                        log_wb = log_excel.Workbooks.Add(1);
                    }
                    else
                    {
                        log_wb = log_excel.Workbooks.Open(Program.logfile);

                    }

                    log_ws = log_wb.Worksheets[1];

                    Range log_xlRange = log_ws.UsedRange;
                    int log_cols = log_xlRange.Columns.Count;
                    int log_rows = log_xlRange.Rows.Count;

                    if (log_cols != 1)
                    {
                        Range range = log_ws.UsedRange;
                        values = (object[,])range.Value2;

                        for (int i = 1; i <= log_rows; i++)
                        {
                            for (int j = 1; j <= log_cols; j++)
                            {
                                if (values[i, j] != null)
                                    logfilearray[i - 1, j - 1] = values[i, j].ToString();
                                else logfilearray[i - 1, j - 1] = "";
                            }
                        }
                    }

                    if (log_cols == 1)
                    {
                        logfilearray[0, 0] = "File Name";
                        logfilearray[0, 1] = "Status";
                        logfilearray[0, 2] = "Date";
                        logfilearray[0, 3] = "Time";
                    }

                    string fileName = System.IO.Path.GetFileName(Program.sourcepath);
                    extension = Path.GetExtension(fileName);
                    fname = Path.GetFileNameWithoutExtension(fileName);

                    logfilearray[log_rows, 0] = fname + extension;
                    logfilearray[log_rows, 1] = "Successful";
                    logfilearray[log_rows, 2] = DateTime.Now.ToString("yyyy-MM-dd");
                    logfilearray[log_rows, 3] = DateTime.Now.ToString("HH:mm:ss");


                    var logstartCell = (Range)log_ws.Cells[1, 1];
                    var logendCell = (Range)log_ws.Cells[log_rows + 1, 4];

                    var logwriteRange = log_ws.Range[logstartCell, logendCell];

                    var columnHeadingsRangeHeader = log_ws.Range[log_ws.Cells[1, 1], log_ws.Cells[1, 4]];
                    columnHeadingsRangeHeader.Interior.Color = 0xD9D9D9;
                    columnHeadingsRangeHeader.Font.Color = XlRgbColor.rgbBlack;
                    columnHeadingsRangeHeader.Font.Size = 12;
                    columnHeadingsRangeHeader.EntireRow.Font.Bold = true;
                    columnHeadingsRangeHeader.Borders.Color = XlRgbColor.rgbBlack;
                    columnHeadingsRangeHeader.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

                    var columnHeadingsRangeData = log_ws.Range[log_ws.Cells[2, 1], log_ws.Cells[log_rows + 1, 4]];
                    columnHeadingsRangeData.Font.Color = XlRgbColor.rgbBlack;
                    columnHeadingsRangeData.Font.Size = 11;
                    columnHeadingsRangeData.Borders.Color = XlRgbColor.rgbBlack;

                    logwriteRange.Value2 = logfilearray;

                    log_wb.SaveAs(Program.logfile);
                    log_wb.Close();
                }
                catch (Exception e)
                {
                    Program.CloseAllExcelAppication(log_ws, log_wb, log_excel);

                    Console.WriteLine(e.Message);
                    Console.WriteLine("Can not Load or Create Master log file.");
                }
                finally
                {
                    Program.CloseAllExcelAppication(log_ws, log_wb, log_excel);
                }

                //auto fit columns
                AutoFitRowColumn fit = new AutoFitRowColumn();
                fit.FittingRowColumn(Program.logfile);


                Console.WriteLine("Source File: " + fname + extension);
                Console.WriteLine("Successfully Processed: " + Program.successful_row + " Rows");
                Console.WriteLine("Failed Processing: " + notsuccessful + " Rows");
                Console.WriteLine("Status: Successfully Processed");
                if(ValidationCheck.logrows!=1)
                {
                    Console.WriteLine("In this source file has some non mandatory field error!");
                    Console.WriteLine("Please check the Error_Log_File.xls in " + @"""Processed_Files\Process Failed\""" + " folder.");
                }
                Console.WriteLine();

                MoveExcelFiles closegr = new MoveExcelFiles();
                if(notsuccessful==0)
                    closegr.MoveSourceExcel();
            }
        }
        public void UnSuccessfulLog()
        {
            int notsuccessful = Source.source_total_rows - Program.successful_row - 3;
            if (notsuccessful != 0 || Program.lengthtoolong == false || Program.continuexecution == false)
            {

                log_excel = new Application();
                log_excel.DisplayAlerts = false;

                try
                {
                    if (!System.IO.File.Exists(Program.logfile))
                    {
                        log_wb = log_excel.Workbooks.Add(1);
                    }
                    else
                    {
                        log_wb = log_excel.Workbooks.Open(Program.logfile);

                    }

                    log_ws = log_wb.Worksheets[1];

                    Range log_xlRange = log_ws.UsedRange;
                    int log_cols = log_xlRange.Columns.Count;
                    int log_rows = log_xlRange.Rows.Count;

                    if (log_cols != 1)
                    {
                        Range range = log_ws.UsedRange;
                        values = (object[,])range.Value2;

                        for (int i = 1; i <= log_rows; i++)
                        {
                            for (int j = 1; j <= log_cols; j++)
                            {
                                if (values[i, j] != null)
                                    logfilearray[i - 1, j - 1] = values[i, j].ToString();
                                else logfilearray[i - 1, j - 1] = "";
                            }
                        }
                    }

                    if (log_cols == 1)
                    {
                        logfilearray[0, 0] = "File Name";
                        logfilearray[0, 1] = "Status";
                        logfilearray[0, 2] = "Date";
                        logfilearray[0, 3] = "Time";
                    }

                    string fileName = System.IO.Path.GetFileName(Program.sourcepath);
                    extension = Path.GetExtension(fileName);
                    fname = Path.GetFileNameWithoutExtension(fileName);

                    logfilearray[log_rows, 0] = fname + extension;
                    logfilearray[log_rows, 1] = "Failed";
                    logfilearray[log_rows, 2] = DateTime.Now.ToString("yyyy-MM-dd");
                    logfilearray[log_rows, 3] = DateTime.Now.ToString("HH:mm:ss");


                    var logstartCell = (Range)log_ws.Cells[1, 1];
                    var logendCell = (Range)log_ws.Cells[log_rows + 1, 4];

                    var logwriteRange = log_ws.Range[logstartCell, logendCell];

                    var columnHeadingsRangeHeader = log_ws.Range[log_ws.Cells[1, 1], log_ws.Cells[1, 4]];
                    columnHeadingsRangeHeader.Interior.Color = 0xD9D9D9;
                    columnHeadingsRangeHeader.Font.Color = XlRgbColor.rgbBlack;
                    columnHeadingsRangeHeader.Font.Size = 12;
                    columnHeadingsRangeHeader.EntireRow.Font.Bold = true;
                    columnHeadingsRangeHeader.Borders.Color = XlRgbColor.rgbBlack;
                    columnHeadingsRangeHeader.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

                    var columnHeadingsRangeData = log_ws.Range[log_ws.Cells[2, 1], log_ws.Cells[log_rows + 1, 4]];
                    columnHeadingsRangeData.Font.Color = XlRgbColor.rgbBlack;
                    columnHeadingsRangeData.Font.Size = 11;
                    columnHeadingsRangeData.Borders.Color = XlRgbColor.rgbBlack;

                    logwriteRange.Value2 = logfilearray;

                    log_wb.SaveAs(Program.logfile);
                    log_wb.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Can not Load or Create the Master log file.");

                    Program.CloseAllExcelAppication(log_ws, log_wb, log_excel);
                }
                finally
                {
                    Program.CloseAllExcelAppication(log_ws, log_wb, log_excel);
                }

                // AutoFit Columns
                AutoFitRowColumn fit = new AutoFitRowColumn();
                fit.FittingRowColumn(Program.logfile);

                if (Program.continuexecution == true)
                {
                    Console.WriteLine("Source File: " + fname + extension);
                    Console.WriteLine("Successfully Processed: " + Program.successful_row + " Rows");
                    Console.WriteLine("Failed Processing: " + notsuccessful + " Rows");
                    if(notsuccessful!=0)
                    {
                        Console.WriteLine("Status: Not Successfully Processed!");
                        Console.WriteLine("Please check the Error_Log_File.xls in " + @"""Processed_Files\Process Failed\""" + " folder.");

                        if (Program.lengthtoolong == false)
                        {
                            string destinationfilenames = fname + extension;
                            string destFile = System.IO.Path.Combine(Program.unsuccessful, destinationfilenames);

                            try
                            {
                                System.IO.File.Move(Program.sourcepath, destFile);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }
                        }

                        Program.lengthtoolong = true;
                    } 
                    Console.WriteLine();
                }
            }
        }
        public void ErrorLog()
        {
            ValidationCheck.errorlogfilearray[0, 0] = "Error Row Number";
            ValidationCheck.errorlogfilearray[0, 1] = "Error Column Number";
            ValidationCheck.errorlogfilearray[0, 2] = "Error Description";

            error_log_excel = new Application();
            error_log_excel.DisplayAlerts = false;

            try
            {
                error_log_wb = error_log_excel.Workbooks.Add(1);

                error_log_ws = error_log_wb.Worksheets[1];

                Range log_xlRange = error_log_ws.UsedRange;
                int log_cols = log_xlRange.Columns.Count;
                int log_rows = log_xlRange.Rows.Count;

                string fileName = System.IO.Path.GetFileName(Program.sourcepath);
                extension = Path.GetExtension(fileName);
                fname = Path.GetFileNameWithoutExtension(fileName);

                string error_folder = fname + DateTime.Now.ToString("yyyyMMddHHmmss");

                destFile = Program.unsuccessful + @"\" + error_folder;

                if (!System.IO.Directory.Exists(destFile))
                {
                    System.IO.Directory.CreateDirectory(destFile);
                }

                var logstartCell = (Range)error_log_ws.Cells[1, 1];
                var logendCell = (Range)error_log_ws.Cells[ValidationCheck.logrows, 3];

                var logwriteRange = error_log_ws.Range[logstartCell, logendCell];

                var columnHeadingsRangeHeader = error_log_ws.Range[error_log_ws.Cells[1, 1], error_log_ws.Cells[1, 3]];
                columnHeadingsRangeHeader.Interior.Color = 0xD9D9D9;
                columnHeadingsRangeHeader.Font.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeHeader.Font.Size = 12;
                columnHeadingsRangeHeader.EntireRow.Font.Bold = true;
                columnHeadingsRangeHeader.Borders.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeHeader.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

                var columnHeadingsRangeData = error_log_ws.Range[error_log_ws.Cells[2, 1], error_log_ws.Cells[ValidationCheck.logrows, 3]];
                columnHeadingsRangeData.Font.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeData.Font.Size = 11;
                columnHeadingsRangeData.Borders.Color = XlRgbColor.rgbBlack;

                logwriteRange.Value2 = ValidationCheck.errorlogfilearray;
                savefilenameas = destFile + @"\" + "Error_Log_File.xls";
                error_log_wb.SaveAs(savefilenameas);

                error_log_wb.Close(0);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Can not Load or Create the Error log file.");
                Program.CloseAllExcelAppication(error_log_ws, error_log_wb, error_log_excel);
            }
            finally
            {
                Program.CloseAllExcelAppication(error_log_ws, error_log_wb, error_log_excel);

            }

            // AutoFit Columns
            AutoFitRowColumn fit = new AutoFitRowColumn();
            string fitfile = destFile + @"\" + "Format_Error_Log_File.xls";
            fit.FittingRowColumn(savefilenameas);

            //// move source file to error folder
            MoveExcelFiles move = new MoveExcelFiles();

            int notsuccessful = Source.source_total_rows - Program.successful_row - 3;

            if (notsuccessful!=0)
                move.MoveFiles(destFile, fname, extension);
        }
        public void FormatErrorLog()
        {
            try
            {
                header_error_log_excel = new Application();
                header_error_log_excel.DisplayAlerts = false;

                header_error_log_wb = header_error_log_excel.Workbooks.Add(1);

                header_error_log_ws = header_error_log_wb.Worksheets[1];

                Range log_xlRange = header_error_log_ws.UsedRange;
                int log_cols = log_xlRange.Columns.Count;
                int log_rows = log_xlRange.Rows.Count;

                string fileName = System.IO.Path.GetFileName(Program.sourcepath);
                extension = Path.GetExtension(fileName);
                fname = Path.GetFileNameWithoutExtension(fileName);

                string error_folder = fname + DateTime.Now.ToString("yyyyMMddHHmmss");

                destFile = Program.unsuccessful + @"\" + error_folder;

                if (!System.IO.Directory.Exists(destFile))
                {
                    System.IO.Directory.CreateDirectory(destFile);
                }

                var logstartCell = (Range)header_error_log_ws.Cells[1, 1];
                var logendCell = (Range)header_error_log_ws.Cells[Source.headerformaterror_rownumber + 1, 3];

                var logwriteRange = header_error_log_ws.Range[logstartCell, logendCell];

                var columnHeadingsRangeHeader = header_error_log_ws.Range[header_error_log_ws.Cells[1, 1], header_error_log_ws.Cells[1, 3]];
                columnHeadingsRangeHeader.Interior.Color = 0xD9D9D9;
                columnHeadingsRangeHeader.Font.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeHeader.Font.Size = 12;
                columnHeadingsRangeHeader.EntireRow.Font.Bold = true;
                columnHeadingsRangeHeader.Borders.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeHeader.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

                var columnHeadingsRangeData = header_error_log_ws.Range[header_error_log_ws.Cells[2, 1], header_error_log_ws.Cells[Source.headerformaterror_rownumber, 3]];
                columnHeadingsRangeData.Font.Color = XlRgbColor.rgbBlack;
                columnHeadingsRangeData.Font.Size = 11;
                columnHeadingsRangeData.Borders.Color = XlRgbColor.rgbBlack;

                logwriteRange.Value2 = Source.headerformaterror;

                savefilenameas = destFile + @"\" + "Format_Error_Log_File.xls";
                header_error_log_wb.SaveAs(savefilenameas);

                header_error_log_wb.Close(0);
            }
            catch (Exception e)
            {
                Console.WriteLine("File directory is too long, directory can't be more than 218 characters.");
                Console.WriteLine("Can not Load or Create the Header Error log file.");
                Program.CloseAllExcelAppication(header_error_log_ws, header_error_log_wb, header_error_log_excel);
            }
            finally
            {
                Program.CloseAllExcelAppication(header_error_log_ws, header_error_log_wb, header_error_log_excel);
            }

            // AutoFit Columns
            AutoFitRowColumn fit = new AutoFitRowColumn();
            fit.FittingRowColumn(savefilenameas);

            //MoveExcelFiles 
            MoveExcelFiles move = new MoveExcelFiles();
            move.MoveFiles(destFile, fname, extension);
        }
    }
}
