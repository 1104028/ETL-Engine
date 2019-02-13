using System;
using System.Globalization;

namespace HRS_ETL_Tool
{
    class DataMerge
    {
        public static int destrow = 0;
        int source_row, rowforbd = 1;
        public static int tracing = 0;
        public static string[,] logarray = new string[2000, 5];
        string datefield;
        int savedestrow;
        public void SaveOutputData()
        {
            string destinationpath = Program.destinationfolder + "Output_ " + Program.fileno + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

            if (destinationpath.Length <= 218)
            {
                ValidationCheck error = new ValidationCheck();

                ValidationCheck.logrows = 1;

                for (var l = 0; l <= 10000; l++)
                {
                    for (var j = 0; j <= 2; j++)
                    {
                        ValidationCheck.errorlogfilearray[l, j] = "";
                    }
                }

                for (source_row = 4; source_row <= Source.source_total_rows; source_row++)
                {
                    if (error.CheckValidity(source_row))
                    {
                        Program.successful_row++;
                        savedestrow = destrow;

                        if (IfSourceColumnExists(Source.Map[0]))
                            Format.destinationfile[destrow, 0] = Source.sourcefile[source_row, Source.Map[0]];

                        Format.destinationfile[destrow, 1] = Source.sourcefile[source_row, Source.Map[1]];
                        Format.destinationfile[destrow, 2] = "";
                        Format.destinationfile[destrow, 3] = Source.sourcefile[source_row, Source.Map[2]] + " " + Source.sourcefile[source_row, Source.Map[3]];
                        if (IfSourceColumnExists(Source.Map[4]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[4]] != "")
                                Format.destinationfile[destrow, 4] = Source.sourcefile[source_row, Source.Map[4]];
                            else
                                Format.destinationfile[destrow, 4] = "-";
                        }
                        else
                            Format.destinationfile[destrow, 4] = "";

                        Format.destinationfile[destrow, 5] = Source.sourcefile[source_row, Source.Map[5]];
                        Format.destinationfile[destrow, 6] = Source.sourcefile[source_row, Source.Map[6]];

                        if (IfSourceColumnExists(Source.Map[7]) && IfSourceColumnExists(Source.Map[8]) && IfSourceColumnExists(Source.Map[9]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[7]] != "" && Source.sourcefile[source_row, Source.Map[8]] != "" && Source.sourcefile[source_row, Source.Map[9]] != "")
                                Format.destinationfile[destrow, 7] = "+" + Source.sourcefile[source_row, Source.Map[7]] + " " + Source.sourcefile[source_row, Source.Map[8]] + Source.sourcefile[source_row, Source.Map[9]];
                            else
                                Format.destinationfile[destrow, 7] = "";
                        }
                        if (IfSourceColumnExists(Source.Map[10]) && IfSourceColumnExists(Source.Map[11]) && IfSourceColumnExists(Source.Map[12]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[10]] != "" && Source.sourcefile[source_row, Source.Map[11]] != "" && Source.sourcefile[source_row, Source.Map[12]] != "")
                            {
                                Format.destinationfile[destrow, 8] = "+" + Source.sourcefile[source_row, Source.Map[10]] + " " + Source.sourcefile[source_row, Source.Map[11]] + Source.sourcefile[source_row, Source.Map[12]];
                            }
                            else
                                Format.destinationfile[destrow, 8] = "";
                        }
                        if (IfSourceColumnExists(Source.Map[13]))
                            Format.destinationfile[destrow, 9] = Source.sourcefile[source_row, Source.Map[13]];
                        if (IfSourceColumnExists(Source.Map[14]))
                            Format.destinationfile[destrow, 10] = Source.sourcefile[source_row, Source.Map[14]];
                        if (IfSourceColumnExists(Source.Map[15]) && IfSourceColumnExists(Source.Map[16]) && IfSourceColumnExists(Source.Map[17]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[15]] != "" && Source.sourcefile[source_row, Source.Map[16]] != "" && Source.sourcefile[source_row, Source.Map[17]] != "")
                            {
                                Format.destinationfile[destrow, 11] = "+" + Source.sourcefile[source_row, Source.Map[15]] + " " + Source.sourcefile[source_row, Source.Map[16]] + Source.sourcefile[source_row, Source.Map[17]];
                            }
                            else
                                Format.destinationfile[destrow, 11] = "";
                        }

                        //Format.destinationfile[destrow, 17] = "";
                        Format.destinationfile[destrow, 20] = "";

                        if (Source.sourcefile[source_row, Source.Map[178]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                            Format.destinationfile[destrow, 21] = "yes";
                        else
                            Format.destinationfile[destrow, 21] = "no";

                        if (IfSourceColumnExists(Source.Map[179]))
                            DoubleConvertion(22, Source.Map[179]);

                        if (IfSourceColumnExists(Source.Map[180]))
                            Format.destinationfile[destrow, 23] = Source.sourcefile[source_row, Source.Map[180]];

                        Format.destinationfile[destrow, 24] = "no";

                        if (IfSourceColumnExists(Source.Map[181]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[181]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 25] = "yes";
                            else
                                Format.destinationfile[destrow, 25] = "no";
                        }

                        if (IfSourceColumnExists(Source.Map[182]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[182]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 26] = "included";
                            else
                                Format.destinationfile[destrow, 26] = "";
                        }

                        if (IfSourceColumnExists(Source.Map[183]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[183]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 27] = "included";
                            else
                                Format.destinationfile[destrow, 27] = "";
                        }

                        if (IfSourceColumnExists(Source.Map[184]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[184]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 28] = "included";
                            else
                                Format.destinationfile[destrow, 28] = "";
                        }

                        if (IfSourceColumnExists(Source.Map[185]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[185]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 29] = "included";
                            else
                                Format.destinationfile[destrow, 29] = "";
                        }

                        if (IfSourceColumnExists(Source.Map[186]))
                        {

                            if (Source.sourcefile[source_row, Source.Map[186]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 30] = "included";
                            else
                                Format.destinationfile[destrow, 30] = "";
                        }

                        if (IfSourceColumnExists(Source.Map[187]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[187]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 31] = "included";
                            else
                                Format.destinationfile[destrow, 31] = "";
                        }

                        Format.destinationfile[destrow, 32] = "";
                        Format.destinationfile[destrow, 33] = "";

                        try
                        {
                            if (IfSourceColumnExists(Source.Map[188]))
                            {
                                string timeanddays = Source.sourcefile[source_row, Source.Map[188]];

                                int leng = timeanddays.Length;

                                if (timeanddays != "" && (timeanddays[leng - 1].Equals(('S')) || timeanddays[leng - 1].Equals(('0'))))
                                {
                                    if (timeanddays[leng - 1].Equals(('S')))
                                    {
                                        Format.destinationfile[destrow, 35] = "";

                                        string[] split = timeanddays.Split('H');
                                        int days = 0;
                                        Int32.TryParse(split[0], out days);

                                        Format.destinationfile[destrow, 34] = (days / 24).ToString();
                                    }
                                    else
                                    {
                                        string time = Source.sourcefile[source_row, Source.Map[188]];

                                        Char delimiter = ':';
                                        string[] split = time.Split(delimiter);

                                        int hour = 0;
                                        Int32.TryParse(split[0], out hour);

                                        int minute = 0;
                                        Int32.TryParse(split[1], out minute);
                                        string output;

                                        if (hour > 12)
                                        {
                                            hour = hour - 12;
                                            if (hour < 10 && minute < 10)
                                            {
                                                output = "0" + hour + ":" + minute + "0" + " p.m.";
                                            }
                                            else if (hour < 10 && minute >= 10)
                                            {
                                                output = "0" + hour + ":" + minute + " p.m.";
                                            }
                                            else if (hour >= 10 && minute < 10)
                                            {
                                                output = hour + ":" + minute + "0" + " p.m.";
                                            }
                                            else
                                            {
                                                output = hour + ":" + minute + " p.m.";
                                            }
                                        }
                                        else
                                        {
                                            if (hour < 10 && minute < 10)
                                            {
                                                output = "0" + hour + ":" + minute + "0" + " a.m.";
                                            }
                                            else if (hour < 10 && minute >= 10)
                                            {
                                                output = "0" + hour + ":" + minute + " a.m.";
                                            }
                                            else if (hour >= 10 && minute < 10)
                                            {
                                                output = hour + ":" + minute + "0" + " a.m.";
                                            }
                                            else
                                            {
                                                output = hour + ":" + minute + " a.m.";
                                            }
                                        }

                                        Format.destinationfile[destrow, 35] = output;
                                        Format.destinationfile[destrow, 34] = "";
                                    }
                                }
                                else
                                {
                                    Format.destinationfile[destrow, 35] = "";
                                    Format.destinationfile[destrow, 34] = "";
                                }

                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        if (IfSourceColumnExists(Source.Map[189]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[189]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 37] = "yes";
                            else
                                Format.destinationfile[destrow, 37] = "no";
                        }

                        if (IfSourceColumnExists(Source.Map[190]))
                        {
                            DoubleConvertion(38, Source.Map[190]);
                        }

                        if (IfSourceColumnExists(Source.Map[191]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[191]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 39] = "yes";
                            else
                                Format.destinationfile[destrow, 39] = "no";
                        }

                        if (IfSourceColumnExists(Source.Map[192]))
                        {
                            DoubleConvertion(40, Source.Map[192]);
                        }

                        if (IfSourceColumnExists(Source.Map[193]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[193]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 41] = "yes";
                            else
                                Format.destinationfile[destrow, 41] = "no";
                        }

                        if (IfSourceColumnExists(Source.Map[194]))
                        {
                            DoubleConvertion(42, Source.Map[194]);
                        }

                        Format.destinationfile[destrow, 43] = "";
                        Format.destinationfile[destrow, 44] = "";

                        if (IfSourceColumnExists(Source.Map[195]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[195]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 45] = "yes";
                            else
                                Format.destinationfile[destrow, 45] = "no";
                        }

                        if (IfSourceColumnExists(Source.Map[196]))
                        {
                            Format.destinationfile[destrow, 46] = Source.sourcefile[source_row, Source.Map[196]];
                        }

                        if (IfSourceColumnExists(Source.Map[197]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[197]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 47] = "yes";
                            else
                                Format.destinationfile[destrow, 47] = "no";
                        }

                        if (IfSourceColumnExists(Source.Map[198]))
                        {
                            if (Source.sourcefile[source_row, Source.Map[198]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                Format.destinationfile[destrow, 49] = "yes";
                            else
                                Format.destinationfile[destrow, 49] = "no";
                        }

                        if (IfSourceColumnExists(Source.Map[199]))
                        {
                            Format.destinationfile[destrow, 50] = Source.sourcefile[source_row, Source.Map[199]];
                        }

                        //transformation
                        //season 1
                        if (Source.sourcefile[source_row, Source.Map[18]] != "" && Source.sourcefile[source_row, Source.Map[19]] != "")
                        {
                            if (CheckSGLDBL(source_row, 20, 21))
                            {
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[18]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[19]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[20]);
                                DoubleConvertion(19, Source.Map[21]);
                                destrow++;
                            }

                            //Creating another row for RT2
                            if (CheckSGLDBL(source_row, 22, 23))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[18]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[19]]; ;
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[22]);
                                DoubleConvertion(19, Source.Map[23]);
                                destrow++;
                            }

                            //Creating another row for RT3
                            if (CheckSGLDBL(source_row, 24, 25))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[18]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[19]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[24]);
                                DoubleConvertion(19, Source.Map[25]);
                                destrow++;
                            }

                            //Creating another row for RT1 & NLRA

                            if (CheckSGLDBL(source_row, 26, 27))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[18]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[19]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[26]);
                                DoubleConvertion(19, Source.Map[27]);
                                destrow++;
                            }

                            //Creating another row for RT2 & NLRA
                            if (CheckSGLDBL(source_row, 28, 29))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[18]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[19]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[28]);
                                DoubleConvertion(19, Source.Map[29]);
                                destrow++;
                            }
                            //Creating another row for RT3 & NLRA
                            if (CheckSGLDBL(source_row, 30, 31))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[18]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[19]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[30]);
                                DoubleConvertion(19, Source.Map[31]);
                                destrow++;
                            }
                        }

                        //sesson 2
                        if (Source.sourcefile[source_row, Source.Map[32]] != "" && Source.sourcefile[source_row, Source.Map[33]] != "")
                        {
                            if (CheckSGLDBL(source_row, 34, 35))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT1 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[32]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[33]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[34]);
                                DoubleConvertion(19, Source.Map[35]);
                                destrow++;
                            }

                            //Creating another row for RT2
                            if (CheckSGLDBL(source_row, 36, 37))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[32]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[33]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[36]);
                                DoubleConvertion(19, Source.Map[37]);
                                destrow++;
                            }

                            //Creating another row for RT3
                            if (CheckSGLDBL(source_row, 38, 39))
                            {
                                CheckRow();
                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[32]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[33]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[38]);
                                DoubleConvertion(19, Source.Map[39]);
                                destrow++;
                            }

                            //Creating another row for RT1 & NLRA
                            if (CheckSGLDBL(source_row, 40, 41))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[32]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[33]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[40]);
                                DoubleConvertion(19, Source.Map[41]);
                                destrow++;
                            }

                            //Creating another row for RT2 & NLRA
                            if (CheckSGLDBL(source_row, 42, 43))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[32]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[33]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[42]);
                                DoubleConvertion(19, Source.Map[43]);
                                destrow++;
                            }
                            //Creating another row for RT3 & NLRA
                            if (CheckSGLDBL(source_row, 44, 45))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[32]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[33]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[44]);
                                DoubleConvertion(19, Source.Map[45]);
                                destrow++;
                            }
                        }
                        //sesson 3
                        if (Source.sourcefile[source_row, Source.Map[46]] != "" && Source.sourcefile[source_row, Source.Map[47]] != "")
                        {
                            if (CheckSGLDBL(source_row, 48, 49))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT1 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription


                                datefield = Source.sourcefile[source_row, Source.Map[46]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[47]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[48]);
                                DoubleConvertion(19, Source.Map[49]);
                                destrow++;
                            }

                            //Creating another row for RT2
                            if (CheckSGLDBL(source_row, 50, 51))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription


                                datefield = Source.sourcefile[source_row, Source.Map[46]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[47]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[50]);
                                DoubleConvertion(19, Source.Map[51]);
                                destrow++;
                            }

                            //Creating another row for RT3
                            if (CheckSGLDBL(source_row, 52, 53))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[46]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[47]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[52]);
                                DoubleConvertion(19, Source.Map[53]);
                                destrow++;
                            }

                            //Creating another row for RT1 & NLRA
                            if (CheckSGLDBL(source_row, 54, 55))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[46]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[47]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[54]);
                                DoubleConvertion(19, Source.Map[55]);
                                destrow++;
                            }

                            //Creating another row for RT2 & NLRA
                            if (CheckSGLDBL(source_row, 56, 57))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[46]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[47]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[56]);
                                DoubleConvertion(19, Source.Map[57]);
                                destrow++;
                            }
                            //Creating another row for RT3 & NLRA
                            if (CheckSGLDBL(source_row, 58, 59))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[46]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[47]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[58]);
                                DoubleConvertion(19, Source.Map[59]);
                                destrow++;
                            }
                        }
                        //sesson 4
                        if (Source.sourcefile[source_row, Source.Map[62]] != "" && Source.sourcefile[source_row, Source.Map[63]] != "")
                        {
                            if (CheckSGLDBL(source_row, 62, 63))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT1 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[60]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[61]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[62]);
                                DoubleConvertion(19, Source.Map[63]);
                                destrow++;
                            }

                            //Creating another row for RT2
                            if (CheckSGLDBL(source_row, 64, 65))
                            {

                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[60]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[61]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[64]);
                                DoubleConvertion(19, Source.Map[65]);
                                destrow++;
                            }

                            //Creating another row for RT3
                            if (CheckSGLDBL(source_row, 66, 67))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[60]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[61]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[66]);
                                DoubleConvertion(19, Source.Map[67]);
                                destrow++;
                            }

                            //Creating another row for RT1 & NLRA
                            if (CheckSGLDBL(source_row, 68, 69))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[60]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[61]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[68]);
                                DoubleConvertion(19, Source.Map[69]);
                                destrow++;
                            }

                            //Creating another row for RT2 & NLRA
                            if (CheckSGLDBL(source_row, 70, 71))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[60]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[61]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[70]);
                                DoubleConvertion(19, Source.Map[71]);
                                destrow++;
                            }
                            //Creating another row for RT3 & NLRA
                            if (CheckSGLDBL(source_row, 72, 73))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[60]];
                                DateConvert(datefield, 15);
                                datefield = Source.sourcefile[source_row, Source.Map[61]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[72]);
                                DoubleConvertion(19, Source.Map[73]);
                                destrow++;
                            }
                        }
                        //sesson 5
                        if (Source.sourcefile[source_row, Source.Map[74]] != "" && Source.sourcefile[source_row, Source.Map[75]] != "")
                        {
                            if (CheckSGLDBL(source_row, 76, 77))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT1 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[74]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[75]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[76]);
                                DoubleConvertion(19, Source.Map[77]);
                                destrow++;
                            }
                            //Creating another row for LRA  RT2
                            if (CheckSGLDBL(source_row, 78, 79))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[74]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[75]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[78]);
                                DoubleConvertion(19, Source.Map[79]);
                                destrow++;
                            }
                            //Creating another row for LRA RT3
                            if (CheckSGLDBL(source_row, 80, 81))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "LRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[74]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[75]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[80]);
                                DoubleConvertion(19, Source.Map[81]);
                                destrow++;
                            }

                            //Creating another row for RT1 & NLRA
                            if (CheckSGLDBL(source_row, 82, 83))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Standard room";//RT2 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[74]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[75]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[82]);
                                DoubleConvertion(19, Source.Map[83]);
                                destrow++;
                            }

                            //Creating another row for RT2 & NLRA
                            if (CheckSGLDBL(source_row, 84, 85))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Superior room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[74]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[75]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[84]);
                                DoubleConvertion(19, Source.Map[85]);
                                destrow++;
                            }
                            //Creating another row for RT3 & NLRA
                            if (CheckSGLDBL(source_row, 86, 87))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    destrowisnotzero();
                                }

                                Format.destinationfile[destrow, 12] = "Business room";//RT3 
                                Format.destinationfile[destrow, 13] = "Corporate Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 14] = "NLRA";//LRA/NLRA
                                Format.destinationfile[destrow, 36] = "";//RateDescription

                                datefield = Source.sourcefile[source_row, Source.Map[74]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[75]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[86]);
                                DoubleConvertion(19, Source.Map[87]);
                                destrow++;
                            }
                        }

                        if (destrow != savedestrow)
                            rowforbd = destrow;
                        else
                            rowforbd = savedestrow + 1;

                        //Creating another row for BD1_RT1                   
                        if (Source.sourcefile[source_row, Source.Map[88]] != "" && Source.sourcefile[source_row, Source.Map[89]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[91]], Source.sourcefile[source_row, Source.Map[92]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1

                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[88]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[89]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[91]);
                                DoubleConvertion(19, Source.Map[92]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[90]];//RateDescription

                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[93]], Source.sourcefile[source_row, Source.Map[94]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1

                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[88]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[89]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, 331];//RateDescription

                                DoubleConvertion(18, Source.Map[93]);
                                DoubleConvertion(19, Source.Map[94]);
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[95]], Source.sourcefile[source_row, Source.Map[96]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[88]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[89]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[95]);
                                DoubleConvertion(19, Source.Map[96]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[90]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[91]], Source.sourcefile[source_row, Source.Map[92]], Source.sourcefile[source_row, Source.Map[93]], Source.sourcefile[source_row, Source.Map[94]], Source.sourcefile[source_row, Source.Map[95]], Source.sourcefile[source_row, Source.Map[96]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[88]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[89]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[90]];//RateDescription
                                destrow++;
                            }

                        }

                        //Creating another row for BD2_RT1
                        if (Source.sourcefile[source_row, Source.Map[97]] != "" && Source.sourcefile[source_row, Source.Map[98]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[100]], Source.sourcefile[source_row, Source.Map[101]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[97]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[98]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[100]);
                                DoubleConvertion(19, Source.Map[101]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[99]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[102]], Source.sourcefile[source_row, Source.Map[103]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[97]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[98]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[102]);
                                DoubleConvertion(19, Source.Map[103]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[99]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[104]], Source.sourcefile[source_row, Source.Map[105]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[97]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[98]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[104]);
                                DoubleConvertion(19, Source.Map[105]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[99]];//RateDescription
                                destrow++;
                            }
                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[100]], Source.sourcefile[source_row, Source.Map[101]], Source.sourcefile[source_row, Source.Map[102]], Source.sourcefile[source_row, Source.Map[103]], Source.sourcefile[source_row, Source.Map[104]], Source.sourcefile[source_row, Source.Map[105]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[97]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[98]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[99]];//RateDescription
                                destrow++;
                            }
                        }

                        //Creating another row for BD3_RT1
                        if (Source.sourcefile[source_row, Source.Map[106]] != "" && Source.sourcefile[source_row, Source.Map[107]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[109]], Source.sourcefile[source_row, Source.Map[110]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";


                                datefield = Source.sourcefile[source_row, Source.Map[106]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[107]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[109]);
                                DoubleConvertion(19, Source.Map[110]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[108]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[111]], Source.sourcefile[source_row, Source.Map[112]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";


                                datefield = Source.sourcefile[source_row, Source.Map[106]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[107]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[111]);
                                DoubleConvertion(19, Source.Map[112]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[108]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[113]], Source.sourcefile[source_row, Source.Map[114]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";


                                datefield = Source.sourcefile[source_row, Source.Map[106]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[107]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[113]);
                                DoubleConvertion(19, Source.Map[114]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[108]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[109]], Source.sourcefile[source_row, Source.Map[110]], Source.sourcefile[source_row, Source.Map[111]], Source.sourcefile[source_row, Source.Map[112]], Source.sourcefile[source_row, Source.Map[113]], Source.sourcefile[source_row, Source.Map[114]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[106]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[107]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[108]];//RateDescription
                                destrow++;
                            }
                        }

                        //Creating another row for BD4_RT1
                        if (Source.sourcefile[source_row, Source.Map[115]] != "" && Source.sourcefile[source_row, Source.Map[116]] != "")
                        {
                            if ((NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[118]], Source.sourcefile[source_row, Source.Map[119]])))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[115]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[116]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[118]);
                                DoubleConvertion(19, Source.Map[119]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[117]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[120]], Source.sourcefile[source_row, Source.Map[121]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";


                                datefield = Source.sourcefile[source_row, Source.Map[115]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[116]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[120]);
                                DoubleConvertion(19, Source.Map[121]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[117]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[122]], Source.sourcefile[source_row, Source.Map[123]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DB
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[115]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[116]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[122]);
                                DoubleConvertion(19, Source.Map[123]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[117]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[118]], Source.sourcefile[source_row, Source.Map[119]], Source.sourcefile[source_row, Source.Map[120]], Source.sourcefile[source_row, Source.Map[121]], Source.sourcefile[source_row, Source.Map[122]], Source.sourcefile[source_row, Source.Map[123]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[115]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[116]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[117]];//RateDescription
                                destrow++;
                            }

                        }
                       
                        //Creating another row for BD5_RT1
                        if (Source.sourcefile[source_row, Source.Map[124]] != "" && Source.sourcefile[source_row, Source.Map[125]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[127]], Source.sourcefile[source_row, Source.Map[128]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[124]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[125]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[127]);
                                DoubleConvertion(19, Source.Map[128]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[126]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[129]], Source.sourcefile[source_row, Source.Map[130]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[124]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[125]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[129]);
                                DoubleConvertion(19, Source.Map[130]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[126]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[131]], Source.sourcefile[source_row, Source.Map[132]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[124]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[125]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[131]);
                                DoubleConvertion(19, Source.Map[132]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[126]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[127]], Source.sourcefile[source_row, Source.Map[128]], Source.sourcefile[source_row, Source.Map[129]], Source.sourcefile[source_row, Source.Map[130]], Source.sourcefile[source_row, Source.Map[131]], Source.sourcefile[source_row, Source.Map[132]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[124]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[125]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[126]];//RateDescription
                                destrow++;
                            }
                        }

                        //Creating another row for BD6_RT1
                        if (Source.sourcefile[source_row, Source.Map[133]] != "" && Source.sourcefile[source_row, Source.Map[134]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[136]], Source.sourcefile[source_row, Source.Map[137]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[133]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[134]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[136]);
                                DoubleConvertion(19, Source.Map[137]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[135]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[138]], Source.sourcefile[source_row, Source.Map[139]]))
                            {

                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[133]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[134]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[138]);
                                DoubleConvertion(19, Source.Map[139]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[135]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[140]], Source.sourcefile[source_row, Source.Map[141]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[133]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[134]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[140]);
                                DoubleConvertion(19, Source.Map[141]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[135]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[136]], Source.sourcefile[source_row, Source.Map[137]], Source.sourcefile[source_row, Source.Map[138]], Source.sourcefile[source_row, Source.Map[139]], Source.sourcefile[source_row, Source.Map[140]], Source.sourcefile[source_row, Source.Map[141]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[133]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[134]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[135]];//RateDescription
                                destrow++;
                            }
                        }

                        //Creating another row for BD7_RT1
                        if (Source.sourcefile[source_row, Source.Map[142]] != "" && Source.sourcefile[source_row, Source.Map[143]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[145]], Source.sourcefile[source_row, Source.Map[146]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[142]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[143]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[145]);
                                DoubleConvertion(19, Source.Map[146]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[144]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[147]], Source.sourcefile[source_row, Source.Map[148]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[142]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[143]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[147]);
                                DoubleConvertion(19, Source.Map[148]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[144]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[149]], Source.sourcefile[source_row, Source.Map[150]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[142]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[143]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[149]);
                                DoubleConvertion(19, Source.Map[150]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[144]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[145]], Source.sourcefile[source_row, Source.Map[146]], Source.sourcefile[source_row, Source.Map[147]], Source.sourcefile[source_row, Source.Map[148]], Source.sourcefile[source_row, Source.Map[149]], Source.sourcefile[source_row, Source.Map[150]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[142]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[143]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[144]];//RateDescription
                                destrow++;
                            }
                        }

                        //Creating another row for BD8_RT1
                        if (Source.sourcefile[source_row, Source.Map[151]] != "" && Source.sourcefile[source_row, Source.Map[152]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[154]], Source.sourcefile[source_row, Source.Map[155]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[151]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[152]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[154]);
                                DoubleConvertion(19, Source.Map[155]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[153]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[156]], Source.sourcefile[source_row, Source.Map[157]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[151]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[152]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[156]);
                                DoubleConvertion(19, Source.Map[157]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[153]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[158]], Source.sourcefile[source_row, Source.Map[159]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[151]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[152]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[158]);
                                DoubleConvertion(19, Source.Map[159]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[153]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[154]], Source.sourcefile[source_row, Source.Map[155]], Source.sourcefile[source_row, Source.Map[156]], Source.sourcefile[source_row, Source.Map[157]], Source.sourcefile[source_row, Source.Map[158]], Source.sourcefile[source_row, Source.Map[159]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[151]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[152]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[153]];//RateDescription
                                destrow++;
                            }
                        }

                        //Creating another row for BD9_RT1
                        if (Source.sourcefile[source_row, Source.Map[160]] != "" && Source.sourcefile[source_row, Source.Map[161]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[163]], Source.sourcefile[source_row, Source.Map[164]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[160]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[161]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[163]);
                                DoubleConvertion(19, Source.Map[164]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[162]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[165]], Source.sourcefile[source_row, Source.Map[166]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }
                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[160]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[161]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[165]);
                                DoubleConvertion(19, Source.Map[166]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[162]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[167]], Source.sourcefile[source_row, Source.Map[168]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";

                                datefield = Source.sourcefile[source_row, Source.Map[160]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[161]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[167]);
                                DoubleConvertion(19, Source.Map[168]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[162]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[163]], Source.sourcefile[source_row, Source.Map[164]], Source.sourcefile[source_row, Source.Map[165]], Source.sourcefile[source_row, Source.Map[166]], Source.sourcefile[source_row, Source.Map[167]], Source.sourcefile[source_row, Source.Map[168]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[160]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[161]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[162]];//RateDescription
                                destrow++;
                            }

                        }

                        //Creating another row for BD10_RT1
                        if (Source.sourcefile[source_row, Source.Map[169]] != "" && Source.sourcefile[source_row, Source.Map[170]] != "")
                        {
                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[172]], Source.sourcefile[source_row, Source.Map[173]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Standard room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";


                                datefield = Source.sourcefile[source_row, Source.Map[169]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[170]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[172]);
                                DoubleConvertion(19, Source.Map[173]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[171]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[174]], Source.sourcefile[source_row, Source.Map[175]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Superior room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";


                                datefield = Source.sourcefile[source_row, Source.Map[169]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[170]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[174]);
                                DoubleConvertion(19, Source.Map[175]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[171]];//RateDescription
                                destrow++;
                            }

                            if (NotEmptyandNumeric(Source.sourcefile[source_row, Source.Map[176]], Source.sourcefile[source_row, Source.Map[177]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                Format.destinationfile[destrow, 13] = "Trade Fair Rate";//Rate Type, LRA_S1_RT1_SGL or DBL
                                Format.destinationfile[destrow, 12] = "Business room";//RT1
                                if (Source.sourcefile[source_row, Source.Map[200]].Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "LRA";
                                else if (Source.sourcefile[source_row, Source.Map[200]].Equals("N", StringComparison.OrdinalIgnoreCase))
                                    Format.destinationfile[destrow, 14] = "NLRA";


                                datefield = Source.sourcefile[source_row, Source.Map[169]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[170]];
                                DateConvert(datefield, 16);

                                DoubleConvertion(18, Source.Map[176]);
                                DoubleConvertion(19, Source.Map[177]);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[171]];//RateDescription
                                destrow++;
                            }

                            if (EmptyandNotNumericBlackoutDate(Source.sourcefile[source_row, Source.Map[172]], Source.sourcefile[source_row, Source.Map[173]], Source.sourcefile[source_row, Source.Map[174]], Source.sourcefile[source_row, Source.Map[175]], Source.sourcefile[source_row, Source.Map[176]], Source.sourcefile[source_row, Source.Map[177]]))
                            {
                                CheckRow();

                                if (destrow != savedestrow)
                                {
                                    LastNLRARow();
                                }

                                //LastNLRARow();

                                BlackOutDate();

                                datefield = Source.sourcefile[source_row, Source.Map[169]];
                                DateConvert(datefield, 15);

                                datefield = Source.sourcefile[source_row, Source.Map[170]];
                                DateConvert(datefield, 16);

                                Format.destinationfile[destrow, 36] = Source.sourcefile[source_row, Source.Map[171]];//RateDescription
                                destrow++;
                            }
                        }
                        CheckRow();
                    }
                }

                if (destrow > 0)
                {
                    OutputFileGenerate outputGe = new OutputFileGenerate();
                    outputGe.WriteOutputExcel();
                }
                if (ValidationCheck.logrows > 1)
                {
                    LogFile xy = new LogFile();
                    xy.ErrorLog();
                }
            }
            else
            {
                Program.lengthtoolong = false;
                Console.WriteLine("Target file directory is too long, directory can't be more than 218 characters.");
            }
        }
        public void DateConvert(string sourcedate, int destcol)
        {
            DateTime outputdater, sourcedater;
            if (!sourcedate.Equals(""))
            {
                try
                {
                    sourcedater = DateTime.Parse(sourcedate);
                    var date = sourcedater.ToString("M/d/yyyy");

                    DateTime.TryParseExact(date.ToString(), "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out outputdater);
                    Format.destinationfile[destrow, destcol] = outputdater.ToString("yyyy-MM-dd");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        public void CopyPreviousRow()
        {
            for (int k = 0; k < 63; k++)
            {
                Format.destinationfile[destrow, k] = Format.destinationfile[destrow - 1, k];
            }
        }
        public void LastNLRARow()
        {
            for (int k = 0; k < 63; k++)
            {
                Format.destinationfile[destrow, k] = Format.destinationfile[rowforbd - 1, k];
            }
        }
        public void CheckRow()
        {
            if (destrow % 898 == 0 && destrow > 0)
            {
                OutputFileGenerate outputg = new OutputFileGenerate();
                if (Program.continuexecution == true)
                {
                    outputg.WriteOutputExcel();
                    destrow = 0;
                }
                else
                {
                    Program.continuexecution = true;
                }
            }
        }
        public void BlackOutDate()
        {
            Format.destinationfile[destrow, 12] = "";
            Format.destinationfile[destrow, 13] = "Blackout Date";
            Format.destinationfile[destrow, 14] = "";
            Format.destinationfile[destrow, 18] = "";
            Format.destinationfile[destrow, 19] = "";
            Format.destinationfile[destrow, 20] = "";
            Format.destinationfile[destrow, 21] = "";
            Format.destinationfile[destrow, 22] = "";
            Format.destinationfile[destrow, 23] = "";
            Format.destinationfile[destrow, 24] = "";

            Format.destinationfile[destrow, 34] = "";
            Format.destinationfile[destrow, 35] = "";
            Format.destinationfile[destrow, 37] = "";
            Format.destinationfile[destrow, 38] = "";
            Format.destinationfile[destrow, 39] = "";

            Format.destinationfile[destrow, 40] = "";
            Format.destinationfile[destrow, 41] = "";
            Format.destinationfile[destrow, 42] = "";
            Format.destinationfile[destrow, 43] = "";
            Format.destinationfile[destrow, 44] = "";
            Format.destinationfile[destrow, 45] = "";

            Format.destinationfile[destrow, 46] = "";
            Format.destinationfile[destrow, 47] = "";
            Format.destinationfile[destrow, 48] = "";
            Format.destinationfile[destrow, 49] = "";
            Format.destinationfile[destrow, 50] = "";
        }
        public void DoubleConvertion(int destcol, int soucecol)
        {
            if (Source.sourcefile[source_row, soucecol] != "" && IsNumeric(Source.sourcefile[source_row, soucecol]) == true)
            {
                double db = Convert.ToDouble(Source.sourcefile[source_row, soucecol]);
                if (db % 1 == 0)
                    Format.destinationfile[destrow, destcol] = (db + ".00").ToString();
                else
                    Format.destinationfile[destrow, destcol] = db.ToString();
            }
            else
            {
                Format.destinationfile[destrow, destcol] = "";
            }
        }
        public bool IsNumeric(String str)
        {
            double d;

            try
            {
                if (str != "")
                    d = Convert.ToDouble(str);
            }
            catch
            {
                return false;
            }
            return true;
        }
        public bool IfSourceColumnExists(int mapvalue)
        {
            if (mapvalue >= 0)
            {
                return true;
            }
            else
                return false;
        }
        public bool NotEmptyandNumeric(string sourceIndex1, string sourceIndex2)
        {
            if ((sourceIndex1 != "" && IsNumeric(sourceIndex1) == true && sourceIndex2 == "") ||
                (sourceIndex1 == "" && sourceIndex2 != "" && IsNumeric(sourceIndex2) == true) ||
                (sourceIndex1 != "" && IsNumeric(sourceIndex1) == true && sourceIndex2 != "" && IsNumeric(sourceIndex2) == true))
                return true;
            else
                return false;
        }
        public bool EmptyandNotNumericBlackoutDate(string sourceIndex1, string sourceIndex2, string sourceIndex3, string sourceIndex4, string sourceIndex5, string sourceIndex6)
        {
            if ((IsNumeric(sourceIndex1) == false && IsNumeric(sourceIndex2) == false) || (IsNumeric(sourceIndex3) == false && IsNumeric(sourceIndex4) == false) || (IsNumeric(sourceIndex5) == false && IsNumeric(sourceIndex6) == false) ||
                (sourceIndex1 == "" && IsNumeric(sourceIndex2) == false) || (sourceIndex3 == "" && IsNumeric(sourceIndex4) == false) || (sourceIndex5 == "" && IsNumeric(sourceIndex6) == false) ||
                (IsNumeric(sourceIndex1) == false && sourceIndex2 == "") || (IsNumeric(sourceIndex3) == false && sourceIndex4 == "") || (IsNumeric(sourceIndex5) == false && sourceIndex6 == "") ||
                (sourceIndex1 == "" && sourceIndex2 == "") || (sourceIndex3 == "" && sourceIndex4 == "") || (sourceIndex5 == "" && sourceIndex6 == ""))
                return true;
            else
                return false;
        }
        public void destrowisnotzero()
        {
            CopyPreviousRow();
        }
        public void destrowisnotzeroforBD()
        {
            LastNLRARow();
        }

        public bool CheckSGLDBL(int source_row, int sourceIndex1, int sourceIndex2)
        {
            if ((Source.sourcefile[source_row, Source.Map[sourceIndex1]] != "" && Source.sourcefile[source_row, Source.Map[sourceIndex2]] == "") ||
                (Source.sourcefile[source_row, Source.Map[sourceIndex1]] == "" && Source.sourcefile[source_row, Source.Map[sourceIndex2]] != "") ||
                ((Source.sourcefile[source_row, Source.Map[sourceIndex1]] != "" && Source.sourcefile[source_row, Source.Map[sourceIndex2]] != "")))
                return true;
            else
                return false;
        }
    }
}
