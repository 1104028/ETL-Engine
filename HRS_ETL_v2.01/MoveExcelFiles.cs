using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace HRS_ETL_Tool
{
    class MoveExcelFiles
    {
        public void MoveFiles(string destFile, string fname, string extension)
        {
            try
            {
                string errorfilegenerate = destFile + @"\";
                string destFilename = System.IO.Path.Combine(errorfilegenerate, errorfilegenerate + fname + extension);

                System.IO.File.Move(Program.sourcepath, destFilename);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        //move excel files to the successful folder
        public void MoveSourceExcel()
        {
            try
            {
                string fileName = System.IO.Path.GetFileName(Program.sourcepath);
                var extension = Path.GetExtension(fileName);
                var fname = Path.GetFileNameWithoutExtension(fileName);

                string destinationfilenames = fname + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extension;
                string destFile = System.IO.Path.Combine(Program.successful, destinationfilenames);

                try
                {
                    System.IO.File.Move(Program.sourcepath, destFile);
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
    }
}
