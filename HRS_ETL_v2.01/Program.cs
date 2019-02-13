using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;

namespace HRS_ETL_Tool
{
    class Program
    {
        public static string sourcepath;// take source path from xml
        public static string fomatpath;// take format path from xml
        public static string logfile, successful, unsuccessful, destination;// declare directories
        public static int fileno, successful_row;
        public static DateTime x;
        public static string destinationfolder;
        public static bool continuexecution = true;
        static string fname, extension;
        public static bool lengthtoolong = true;

        static void Main(string[] args)
        {
            x = DateTime.Now;//store current time

            XmlDocument locationdoc = new XmlDocument();
            string path = System.IO.Directory.GetCurrentDirectory() + "\\config.xml";//store xml path
            if (System.IO.File.Exists(path))
            {
                locationdoc.Load(path);
                XmlNode node = locationdoc.SelectSingleNode("/root/appconfig");

                string sourcedirectory = node["source_directory"].InnerXml;//retrive source directory from xml

                string[] FileNames_array;

                fomatpath = node["format_directory"].InnerXml;//retrive format directory from xml

                string dpath = node["target_directory"].InnerXml;//retrive target directory from xml

                if (System.IO.Directory.Exists(sourcedirectory))
                {
                    FileNames_array = System.IO.Directory.GetFileSystemEntries(sourcedirectory, "*.xls*", System.IO.SearchOption.AllDirectories);//filter source files for excel files
                    if (System.IO.File.Exists(Program.fomatpath))
                    {
                        if (System.IO.Directory.Exists(dpath))
                        {
                            destination = dpath + @"\";//create destination path
                            if (System.IO.Directory.Exists(node["processed_source_directory"].InnerXml))
                            {
                                if (FileNames_array.Length != 0)
                                {
                                    string log = destination;
                                    logfile = log + "Master_Log_File.xls";//create master log file

                                    successful = node["processed_source_directory"].InnerXml + @"\" + "Process Sucessfull";//create successful directory path
                                    if (!System.IO.Directory.Exists(successful))
                                    {
                                        System.IO.Directory.CreateDirectory(successful);//create successful directory
                                    }

                                    unsuccessful = node["processed_source_directory"].InnerXml + @"\" + "Process Failed";//create unsuccessful directory path
                                    if (!System.IO.Directory.Exists(unsuccessful))
                                    {
                                        System.IO.Directory.CreateDirectory(unsuccessful);//create successful directory
                                    }
                                    Console.WriteLine("Processing the source files, please wait .........");
                                    Console.WriteLine();
                                    Format format = new Format();

                                    format.OpenFormatExcel();//call Foramt class

                                    if (Format.formatNumCols == 63 && Format.format_flag == true)
                                    {
                                        for (int i = 0; i < FileNames_array.Length; i++)//process all source files
                                        {
                                            sourcepath = FileNames_array[i];

                                            Source source = new Source();
                                            DataMerge merge = new DataMerge();
                                            LogFile loggenerate = new LogFile();

                                            string fileName = System.IO.Path.GetFileName(sourcepath);//get source file name with extension
                                            extension = Path.GetExtension(fileName);//get source file extension
                                            fname = Path.GetFileNameWithoutExtension(fileName);//get source file name without extension

                                            try
                                            {
                                                source.OpenSourceExcel();// call for open source excel file
                                                destinationfolder = destination + fname + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + @"\";//create destination folder with date time

                                                DataMerge.destrow = 0;// assign destination row which is used in DataMerge class
                                                fileno = 1;// assign output excel file no which is used in OutputFileGenerate class

                                                if (continuexecution == true)// check format or header error occurs or not
                                                {
                                                    try
                                                    {
                                                        successful_row = 0;// assign successful_row value which is then used DataMerge class and Showing Notification
                                                        merge.SaveOutputData();// call DataMerge class for processing source file

                                                        
                                                        int notsuccessful = Source.source_total_rows - successful_row - 3;
                                                        
                                                            try
                                                            {
                                                                if(notsuccessful==0)
                                                                    loggenerate.SuccessfulLog();// call for create successful log excel file
                                                                else
                                                                loggenerate.UnSuccessfulLog();// call for create unsuccessful log excel file                 
                                                            }
                                                            catch (Exception e)
                                                            {
                                                                Console.WriteLine(e.Message);
                                                            }
                                                        
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        Console.WriteLine(e.Message);
                                                    }
                                                }
                                                else
                                                {
                                                    loggenerate.UnSuccessfulLog();// generate unsuccessful log for format or header erros.
                                                    continuexecution = true;// assign true for processing next source file
                                                    UnInterruptExecution();// call for showing notification header or format error file
                                                }
                                            }
                                            catch (Exception e)
                                            {
                                                Console.WriteLine(e.Message);
                                            }
                                        }
                                        SuccessfullyFinishExecution();// call if no error occurs 
                                    }
                                    else
                                    {
                                        Console.WriteLine("Format File Columns are not Valid, Please Select the Format File Which Has 63 Columns.");
                                        DirectoryError();// call if directory error occurs
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Source Dircetory is Empty!");
                                    Console.WriteLine("Please Open the Source Dircetory and Put Source Files.");
                                    DirectoryError();// call if directory error occurs
                                }
                            }
                            else
                            {
                                Console.WriteLine("Processed Directory Not Found!");
                                Console.WriteLine("Please Open the config.xml file and set Processed Directory path.");
                                DirectoryError();// call if directory error occurs
                            }
                        }
                        else
                        {
                            Console.WriteLine("Destination Directory Not Found!");
                            Console.WriteLine("Please Open the config.xml file and set Valid Destination  paths");
                            DirectoryError();// call if directory error occurs
                        }
                    }
                    else
                    {
                        Console.WriteLine("Format Dircetory Not Found!");
                        Console.WriteLine("Please Open the config.xml file and set Valid format paths");
                        DirectoryError();// call if directory error occurs
                    }

                }
                else
                {
                    Console.WriteLine("Source Dircetory is Not Found");
                    Console.WriteLine("Please Open the config file and put Source Dircetory.");
                    DirectoryError();// call if directory error occurs
                }
            }
            else
            {
                Console.WriteLine("Please Read Installation and User Manuals, and Put config.xml in root Folder.");
                DirectoryError();// call if config.xml not found
            }

        }
        public static void SuccessfullyFinishExecution()
        {
            DateTime y = DateTime.Now;// get after processed source file
            Console.WriteLine("Processing time: " + (y - x));// calculate the execution time
            Console.WriteLine("Please press any key to Complete the Process....\n");
            Console.ReadKey();// take a key for finish process
        }
        public static void UnsuccessfullyFinishExecution()
        {
            Console.WriteLine("Source File: " + fname + extension);// print source file name and extension
            DateTime y = DateTime.Now;// get after processed source file
            Console.WriteLine("Status: Not Successfully Processed!");
            Console.WriteLine("Processing time: " + (y - x));
            Console.WriteLine("Please press any key to Complete the Process....\n");
            Console.ReadKey();
            Environment.Exit(0);// exit from the application if in format file has any error
        }
        // In any source files has any type error it will continue processing
        public static void UnInterruptExecution()
        {
            Console.WriteLine("Source File: " + fname + extension);
            Console.WriteLine("Status: Not Successfully Processed!");
            Console.WriteLine("Please check the Format_Error_Log_File.xls in " + @"""Processed_Files\Process Failed\""" + " folder.");
            Console.WriteLine();
        }
        // if file directory is more than 218 character
        public static void FileTooLong()
        {
            Console.WriteLine("Source File: " + fname + extension);
            Console.WriteLine("Status: Not Successfully Processed!");
            Console.WriteLine("Target file directory can't be more than 218 characters.");
            Console.WriteLine();
        }
        //close all oppend or created excel file
        public static void CloseAllExcelAppication(Worksheet ws, Workbook wb, Application excel)
        {
            if (ws != null) Marshal.ReleaseComObject(ws);//close worksheet
            if (wb != null) Marshal.ReleaseComObject(wb);//close workbook
            if (excel != null) excel.Quit();//close excel

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        // execute when directory error occurs
        public static void DirectoryError()
        {
            Console.WriteLine("Please press any key to Complete the Process....\n");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}

