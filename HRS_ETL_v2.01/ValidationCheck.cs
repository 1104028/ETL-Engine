using System;
using System.Globalization;

namespace HRS_ETL_Tool
{
    class ValidationCheck
    {
        string date;
        public static string[,] errorlogfilearray = new string[10001, 3];
        public static int[,] errorrowcolumn = new int[1000, 3];
        public static bool valid = true;
        public static int logrows = 1;
        //int column_number = 0;
        bool lastrmbd;
        public bool CheckValidity(int error_row)
        {
            valid = true;
            lastrmbd = false;
            try
            {
                IsEmptyMapping(error_row, Source.Map[1]);

                IsEmptytransformation(error_row, Source.Map[2], Source.Map[3]);

                IsEmptyMapping(error_row, Source.Map[5]);
                IsEmptyMapping(error_row, Source.Map[6]);

                IsEmptyMapping(error_row, Source.Map[178]);

                IsEmptyMapping(error_row, Source.Map[180]);

                if (!Source.Map[188].Equals(""))
                {
                    CancellationDays(error_row, Source.Map[188]);
                }

                //transformation
                //season 1
                if (Source.sourcefile[error_row, Source.Map[18]] != "" && Source.sourcefile[error_row, Source.Map[19]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[18]];
                    DateConvert(date, error_row, Source.Map[18]);

                    date = Source.sourcefile[error_row, Source.Map[19]];
                    DateConvert(date, error_row, Source.Map[19]);

                    if (CheckSGLDBL(error_row,20,21))
                    {
                        DoubleConvertion(error_row, Source.Map[20]);
                        DoubleConvertion(error_row, Source.Map[21]);
                    }

                    //validation Check for RT2
                    if (CheckSGLDBL(error_row, 22, 23))
                    {
                        DoubleConvertion(error_row, Source.Map[22]);
                        DoubleConvertion(error_row, Source.Map[23]);
                    }

                    //validation Check for RT3
                    if (CheckSGLDBL(error_row, 24, 25))
                    {
                        DoubleConvertion(error_row, Source.Map[24]);
                        DoubleConvertion(error_row, Source.Map[25]);
                    }

                    //validation Check for RT1 & NLRA
                    if (CheckSGLDBL(error_row, 26, 27))
                    {
                        DoubleConvertion(error_row, Source.Map[26]);
                        DoubleConvertion(error_row, Source.Map[27]);
                    }

                    //validation Check for RT2 & NLRA
                    if (CheckSGLDBL(error_row, 28, 29))
                    {
                        DoubleConvertion(error_row, Source.Map[28]);
                        DoubleConvertion(error_row, Source.Map[29]);
                    }

                    //validation Check for RT3 & NLRA
                    if (CheckSGLDBL(error_row, 30, 31))
                    {
                        DoubleConvertion(error_row, Source.Map[30]);
                        DoubleConvertion(error_row, Source.Map[31]);
                    }
                }

                //sesson 2
                if (Source.sourcefile[error_row, Source.Map[32]] != "" && Source.sourcefile[error_row, Source.Map[33]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[32]];
                    DateConvert(date, error_row, Source.Map[32]);

                    date = Source.sourcefile[error_row, Source.Map[33]];
                    DateConvert(date, error_row, Source.Map[33]);

                    if (CheckSGLDBL(error_row, 34, 35))
                    {
                        DoubleConvertion(error_row, Source.Map[34]);
                        DoubleConvertion(error_row, Source.Map[35]);
                    }

                    //validation Check for RT2
                    if (CheckSGLDBL(error_row, 36, 37))
                    {
                        DoubleConvertion(error_row, Source.Map[36]);
                        DoubleConvertion(error_row, Source.Map[37]);
                    }

                    //validation Check for RT3
                    if (CheckSGLDBL(error_row, 38, 39))
                    {
                        DoubleConvertion(error_row, Source.Map[38]);
                        DoubleConvertion(error_row, Source.Map[39]);
                    }

                    //validation Check for RT1 & NLRA
                    if (CheckSGLDBL(error_row, 40, 41))
                    {
                        DoubleConvertion(error_row, Source.Map[40]);
                        DoubleConvertion(error_row, Source.Map[41]);
                    }

                    //validation Check for RT2 & NLRA
                    if (CheckSGLDBL(error_row, 42, 43))
                    {
                        DoubleConvertion(error_row, Source.Map[42]);
                        DoubleConvertion(error_row, Source.Map[43]);
                    }

                    //validation Check for RT3 & NLRA
                    if (CheckSGLDBL(error_row, 44, 45))
                    {
                        DoubleConvertion(error_row, Source.Map[44]);
                        DoubleConvertion(error_row, Source.Map[45]);
                    }
                }
                //sesson 3
                if (Source.sourcefile[error_row, Source.Map[46]] != "" && Source.sourcefile[error_row, Source.Map[47]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[46]];
                    DateConvert(date, error_row, Source.Map[46]);

                    date = Source.sourcefile[error_row, Source.Map[47]];
                    DateConvert(date, error_row, Source.Map[47]);

                    if (CheckSGLDBL(error_row, 48, 49))
                    {
                        DoubleConvertion(error_row, Source.Map[48]);
                        DoubleConvertion(error_row, Source.Map[49]);
                    }

                    //validation Check for RT2

                    if (CheckSGLDBL(error_row, 50, 51))
                    {
                        DoubleConvertion(error_row, Source.Map[50]);
                        DoubleConvertion(error_row, Source.Map[51]);
                    }

                    //validation Check for RT3
                    if (CheckSGLDBL(error_row, 52, 53))
                    {
                        DoubleConvertion(error_row, Source.Map[52]);
                        DoubleConvertion(error_row, Source.Map[53]);
                    }

                    //validation Check for RT1 & NLRA
                    if (CheckSGLDBL(error_row, 54, 55))
                    {
                        DoubleConvertion(error_row, Source.Map[54]);
                        DoubleConvertion(error_row, Source.Map[55]);
                    }

                    //validation Check for RT2 & NLRA
                    if (CheckSGLDBL(error_row, 56, 57))
                    {
                        DoubleConvertion(error_row, Source.Map[56]);
                        DoubleConvertion(error_row, Source.Map[57]);
                    }

                    //validation Check for RT3 & NLRA
                    if (CheckSGLDBL(error_row, 58, 59))
                    {
                        DoubleConvertion(error_row, Source.Map[58]);
                        DoubleConvertion(error_row, Source.Map[59]);
                    }
                }

                //sesson 4
                if (Source.sourcefile[error_row, Source.Map[62]] != "" && Source.sourcefile[error_row, Source.Map[63]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[60]];
                    DateConvert(date, error_row, Source.Map[60]);

                    date = Source.sourcefile[error_row, Source.Map[61]];
                    DateConvert(date, error_row, Source.Map[61]);

                    if (CheckSGLDBL(error_row, 62, 63))
                    {
                        DoubleConvertion(error_row, Source.Map[62]);
                        DoubleConvertion(error_row, Source.Map[63]);
                    }

                    //validation Check for RT2
                    if (CheckSGLDBL(error_row, 64, 65))
                    {
                        DoubleConvertion(error_row, Source.Map[64]);
                        DoubleConvertion(error_row, Source.Map[65]);
                    }

                    //validation Check for RT3
                    if (CheckSGLDBL(error_row, 66, 67))
                    {
                        DoubleConvertion(error_row, Source.Map[66]);
                        DoubleConvertion(error_row, Source.Map[67]);
                    }

                    //validation Check for RT1 & NLRA
                    if (CheckSGLDBL(error_row, 68, 69))
                    {
                        DoubleConvertion(error_row, Source.Map[68]);
                        DoubleConvertion(error_row, Source.Map[69]);
                    }

                    //validation Check for RT2 & NLRA
                    if (CheckSGLDBL(error_row, 70, 71))
                    {
                        DoubleConvertion(error_row, Source.Map[70]);
                        DoubleConvertion(error_row, Source.Map[71]);
                    }

                    //validation Check for RT3 & NLRA
                    if (CheckSGLDBL(error_row, 72, 73))
                    {
                        DoubleConvertion(error_row, Source.Map[72]);
                        DoubleConvertion(error_row, Source.Map[73]);
                    }
                }

                //sesson 5
                if (Source.sourcefile[error_row, Source.Map[74]] != "" && Source.sourcefile[error_row, Source.Map[75]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[74]];
                    DateConvert(date, error_row, Source.Map[74]);

                    date = Source.sourcefile[error_row, Source.Map[75]];
                    DateConvert(date, error_row, Source.Map[75]);

                    if (CheckSGLDBL(error_row, 76, 77))
                    {
                        DoubleConvertion(error_row, Source.Map[76]);
                        DoubleConvertion(error_row, Source.Map[77]);
                    }

                    //validation Check for LRA  RT2
                    if (CheckSGLDBL(error_row, 78, 79))
                    {
                        DoubleConvertion(error_row, Source.Map[78]);
                        DoubleConvertion(error_row, Source.Map[79]);
                    }

                    //validation Check for LRA RT3
                    if (CheckSGLDBL(error_row, 80, 81))
                    {

                        DoubleConvertion(error_row, Source.Map[80]);
                        DoubleConvertion(error_row, Source.Map[81]);
                    }

                    //validation Check for RT1 & NLRA
                    if (CheckSGLDBL(error_row, 82, 83))
                    {

                        DoubleConvertion(error_row, Source.Map[82]);
                        DoubleConvertion(error_row, Source.Map[83]);
                    }

                    //validation Check for RT2 & NLRA
                    if (CheckSGLDBL(error_row, 84, 85))
                    {

                        DoubleConvertion(error_row, Source.Map[84]);
                        DoubleConvertion(error_row, Source.Map[85]);
                    }

                    //validation Check for RT3 & NLRA
                    if (CheckSGLDBL(error_row, 86, 87))
                    {

                        DoubleConvertion(error_row, Source.Map[86]);
                        DoubleConvertion(error_row, Source.Map[87]);
                    }
                }

                //validation Check for BD1_RT1
                if (Source.sourcefile[error_row, Source.Map[88]] != "" && Source.sourcefile[error_row, Source.Map[89]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[88]];
                    DateConvert(date, error_row, Source.Map[88]);

                    date = Source.sourcefile[error_row, Source.Map[89]];
                    DateConvert(date, error_row, Source.Map[89]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[91]], Source.sourcefile[error_row, Source.Map[92]]))
                    {
                        LastarmvilBD(error_row, Source.Map[200]);
                        lastrmbd = true;
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[93]], Source.sourcefile[error_row, Source.Map[94]]))
                    {
                        if(lastrmbd==false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        } 
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[95]], Source.sourcefile[error_row, Source.Map[96]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }
                }
                //validation Check for BD2_RT1
                if (Source.sourcefile[error_row, Source.Map[97]] != "" && Source.sourcefile[error_row, Source.Map[98]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[97]];
                    DateConvert(date, error_row, Source.Map[97]); ;

                    date = Source.sourcefile[error_row, Source.Map[98]];
                    DateConvert(date, error_row, Source.Map[98]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[100]], Source.sourcefile[error_row, Source.Map[101]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[102]], Source.sourcefile[error_row, Source.Map[103]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[104]], Source.sourcefile[error_row, Source.Map[105]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }
                }
                //validation Check for BD3_RT1
                if (Source.sourcefile[error_row, Source.Map[106]] != "" && Source.sourcefile[error_row, Source.Map[107]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[106]];
                    DateConvert(date, error_row, Source.Map[106]);

                    date = Source.sourcefile[error_row, Source.Map[107]];
                    DateConvert(date, error_row, Source.Map[107]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[109]], Source.sourcefile[error_row, Source.Map[110]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[111]], Source.sourcefile[error_row, Source.Map[112]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[113]], Source.sourcefile[error_row, Source.Map[114]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }
                }
                //validation Check for BD4_RT1
                if (Source.sourcefile[error_row, Source.Map[115]] != "" && Source.sourcefile[error_row, Source.Map[116]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[115]];
                    DateConvert(date, error_row, Source.Map[115]);

                    date = Source.sourcefile[error_row, Source.Map[116]];
                    DateConvert(date, error_row, Source.Map[116]);

                    if ((NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[118]], Source.sourcefile[error_row, Source.Map[119]])))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[120]], Source.sourcefile[error_row, Source.Map[121]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[122]], Source.sourcefile[error_row, Source.Map[123]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }
                }
                //validation Check for BD5_RT1
                if (Source.sourcefile[error_row, Source.Map[124]] != "" && Source.sourcefile[error_row, Source.Map[125]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[124]];
                    DateConvert(date, error_row, Source.Map[124]);

                    date = Source.sourcefile[error_row, Source.Map[125]];
                    DateConvert(date, error_row, Source.Map[125]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[127]], Source.sourcefile[error_row, Source.Map[128]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[129]], Source.sourcefile[error_row, Source.Map[130]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        } 
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[131]], Source.sourcefile[error_row, Source.Map[132]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        } 
                    }
                }
                //validation Check for BD6_RT1
                if (Source.sourcefile[error_row, Source.Map[134]] != "" && Source.sourcefile[error_row, Source.Map[135]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[133]];
                    DateConvert(date, error_row, Source.Map[133]);


                    date = Source.sourcefile[error_row, Source.Map[134]];
                    DateConvert(date, error_row, Source.Map[134]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[136]], Source.sourcefile[error_row, Source.Map[137]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[138]], Source.sourcefile[error_row, Source.Map[139]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[140]], Source.sourcefile[error_row, Source.Map[141]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        } 
                    }
                }
                //validation Check for BD7_RT1
                if (Source.sourcefile[error_row, Source.Map[142]] != "" && Source.sourcefile[error_row, Source.Map[143]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[142]];
                    DateConvert(date, error_row, Source.Map[142]);

                    date = Source.sourcefile[error_row, Source.Map[143]];
                    DateConvert(date, error_row, Source.Map[143]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[145]], Source.sourcefile[error_row, Source.Map[146]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }  
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[147]], Source.sourcefile[error_row, Source.Map[148]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        } 
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[149]], Source.sourcefile[error_row, Source.Map[150]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }
                }

                //validation Check for BD8_RT1
                if (Source.sourcefile[error_row, Source.Map[151]] != "" && Source.sourcefile[error_row, Source.Map[152]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[151]];
                    DateConvert(date, error_row, Source.Map[151]);

                    date = Source.sourcefile[error_row, Source.Map[152]];
                    DateConvert(date, error_row, Source.Map[152]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[154]], Source.sourcefile[error_row, Source.Map[155]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        } 
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[156]], Source.sourcefile[error_row, Source.Map[157]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[158]], Source.sourcefile[error_row, Source.Map[159]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }
                }

                //validation Check for BD9_RT1
                if (Source.sourcefile[error_row, Source.Map[160]] != "" && Source.sourcefile[error_row, Source.Map[161]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[160]];
                    DateConvert(date, error_row, Source.Map[160]);

                    date = Source.sourcefile[error_row, Source.Map[161]];
                    DateConvert(date, error_row, Source.Map[161]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[163]], Source.sourcefile[error_row, Source.Map[164]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[165]], Source.sourcefile[error_row, Source.Map[166]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        } 
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[167]], Source.sourcefile[error_row, Source.Map[168]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }  
                    }
                }

                //validation Check for BD10_RT1
                if (Source.sourcefile[error_row, Source.Map[169]] != "" && Source.sourcefile[error_row, Source.Map[170]] != "")
                {
                    date = Source.sourcefile[error_row, Source.Map[169]];
                    DateConvert(date, error_row, Source.Map[169]);

                    date = Source.sourcefile[error_row, Source.Map[170]];
                    DateConvert(date, error_row, Source.Map[170]);

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[172]], Source.sourcefile[error_row, Source.Map[173]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[174]], Source.sourcefile[error_row, Source.Map[175]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }
                    }

                    if (NotEmptyandNumeric(Source.sourcefile[error_row, Source.Map[176]], Source.sourcefile[error_row, Source.Map[177]]))
                    {
                        if (lastrmbd == false)
                        {
                            LastarmvilBD(error_row, Source.Map[200]);
                            lastrmbd = true;
                        }   
                    }
                }
            }
            catch
            { 
                valid = false;
            }

            return valid;
        }
             
        public void DateConvert(string sourcedate, int error_row, int colnum)
        {
            if (sourcedate.Equals(""))
            {
                valid = false;
                string colletter = GetExcelColumnName(colnum);
                errorlogfilearray[logrows, 0] = error_row.ToString();
                errorlogfilearray[logrows, 1] = colletter;
                errorlogfilearray[logrows, 2] = "Date Field Value Can't be empty because start or end has date.";
                logrows++;
            }
            else if (Parsedate(sourcedate) == false)
            {
                valid = false;
                string colletter = GetExcelColumnName(colnum);
                errorlogfilearray[logrows, 0] = error_row.ToString();
                errorlogfilearray[logrows, 1] = colletter;
                errorlogfilearray[logrows, 2] = "Date Field format is not valid.";
                logrows++;
            }
        }
        public void DoubleConvertion(int error_row, int sourcecol1)
        {
            if (IsNumeric(Source.sourcefile[error_row, sourcecol1]) == false)
            {
                string colletter = GetExcelColumnName(sourcecol1);
                errorlogfilearray[logrows, 0] = error_row.ToString();
                errorlogfilearray[logrows, 1] = colletter;
                errorlogfilearray[logrows, 2] = "Data is not Numeric.";
                logrows++;
                valid = false;
            }
        }
        public void EmptyForLRA(int error_row, int sourcecol1, int sourcecol2)
        {
            string colletter = GetExcelColumnName(sourcecol1);
            string colletter1 = GetExcelColumnName(sourcecol2);

            errorlogfilearray[logrows, 0] = error_row.ToString();
            errorlogfilearray[logrows, 1] = colletter + "," + colletter1;
            errorlogfilearray[logrows, 2] = "Data can't be Empty at a time for both rows.";
            logrows++;
            valid = false;

        }
        public void IsEmptyMapping(int error_row, int sourcecol)
        {
            if (Source.sourcefile[error_row, sourcecol] == "")
            {
                string colletter = GetExcelColumnName(sourcecol);

                errorlogfilearray[logrows, 0] = error_row.ToString();
                errorlogfilearray[logrows, 1] = colletter;
                errorlogfilearray[logrows, 2] = "This is a Mandatory Field for Target File, Data Cannot be Empty.";
                logrows++;
                valid = false;
            }

        }
        public void IsEmptytransformation(int error_row, int sourcecol1, int sourcecol2)
        {
            if (Source.sourcefile[error_row, sourcecol1] == "" && Source.sourcefile[error_row, sourcecol2] == "")
            {
                string colletter = GetExcelColumnName(sourcecol1);
                string colletter1 = GetExcelColumnName(sourcecol2);

                errorlogfilearray[logrows, 0] = error_row.ToString();
                errorlogfilearray[logrows, 1] = colletter + "," + colletter1;
                errorlogfilearray[logrows, 2] = "This is a Mandatory Field for Target File, Data Cannot be Empty.";
                logrows++;
                valid = false;
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
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        public bool Parsedate(string sourcedate)
        {
            DateTime outputdater, sourcedater;
            try
            {
                sourcedater = DateTime.Parse(sourcedate);
                var date = sourcedater.ToString("M/d/yyyy");

                DateTime.TryParseExact(date.ToString(), "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out outputdater);
                string x = outputdater.ToString("yyyy-MM-dd");

                return true;
            }
            catch
            {
                return false;
            }
        }
        public void CancellationDays(int error_row, int sourcecol)
        {
            string timeanddays = Source.sourcefile[error_row, Source.Map[188]];

            int leng = timeanddays.Length;

            if (timeanddays != "" && (timeanddays[leng - 1].Equals(('S')) || timeanddays[leng - 1].Equals(('0'))))
            {

            }
            else
            {
                string colletter = GetExcelColumnName(sourcecol);

                errorlogfilearray[logrows, 0] = error_row.ToString();
                errorlogfilearray[logrows, 1] = colletter;
                errorlogfilearray[logrows, 2] = "Free Cancellation days or time is not valid, please follow " + @"""18:00 or 24HRS """ + "format";
                logrows++;
            }
        }
        public bool EmptyandNotNumericBlackoutDate(int error_row,int sourceIndex1, int sourceIndex2, int sourceIndex3, int sourceIndex4, int sourceIndex5, int sourceIndex6)
        {
            if ((Source.sourcefile[error_row, Source.Map[sourceIndex1]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex2]] == "") || (Source.sourcefile[error_row, Source.Map[sourceIndex1]] == "" && Source.sourcefile[error_row, Source.Map[sourceIndex2]] != "") || (Source.sourcefile[error_row, Source.Map[sourceIndex1]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex2]] != "") ||
                (Source.sourcefile[error_row, Source.Map[sourceIndex3]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex4]] == "") || (Source.sourcefile[error_row, Source.Map[sourceIndex3]] == "" && Source.sourcefile[error_row, Source.Map[sourceIndex4]] != "") || (Source.sourcefile[error_row, Source.Map[sourceIndex3]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex4]] != "") ||
                (Source.sourcefile[error_row, Source.Map[sourceIndex5]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex6]] == "") || (Source.sourcefile[error_row, Source.Map[sourceIndex5]] == "" && Source.sourcefile[error_row, Source.Map[sourceIndex6]] != "") || (Source.sourcefile[error_row, Source.Map[sourceIndex5]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex6]] != ""))
            {
                return true;
            }
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
        public void LastarmvilBD(int error_row, int sourcecol)
        {
            if (Source.sourcefile[error_row, sourcecol] == "" || (!Source.sourcefile[error_row, sourcecol].Equals("Y", StringComparison.OrdinalIgnoreCase) && !Source.sourcefile[error_row, sourcecol].Equals("N", StringComparison.OrdinalIgnoreCase)))
            { 
                string colletter = GetExcelColumnName(sourcecol);

                errorlogfilearray[logrows, 0] = error_row.ToString();
                errorlogfilearray[logrows, 1] = colletter;
                errorlogfilearray[logrows, 2] = "This is a Mandatory Field for Target File, Data Cannot be Empty and value can be Y/y, N/n.";
                logrows++;
                valid = false;
            }
        }
        public bool CheckSGLDBL(int error_row, int sourceIndex1,int sourceIndex2)
        {
            if ((Source.sourcefile[error_row, Source.Map[sourceIndex1]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex2]] == "") ||
                (Source.sourcefile[error_row, Source.Map[sourceIndex1]] == "" && Source.sourcefile[error_row, Source.Map[sourceIndex2]] != "") ||
                ((Source.sourcefile[error_row, Source.Map[sourceIndex1]] != "" && Source.sourcefile[error_row, Source.Map[sourceIndex2]] != "")))
                return true;
            else
                return false;
        }
    }
}

