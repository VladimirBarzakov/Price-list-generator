using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;




namespace ReffPrice
{
    class Program
    {
        class Machine                                                                               // Клас Машина
        {
            public string OldPosition { get; set; }
            public string NewPosition { get; set; }
            public int? RefNum { get; set; }
            public string Type { get; set; }
            public string Brand { get; set; }
            public string Model { get; set; }
            public string Clarification { get; set; }
            public double? NominalHP { get; set; }
            public double? OldPrice { get; set; }
            public double? NewPrice { get; set; }
            public double? AlowedPrice { get; set; }
            public double? ReffPrice { get; set; }
            public string Provider { get; set; }
            public string DateFrom { get; set; }
            public string DateTo { get; set; }
            public double? Inflation { get; set; }
            public string CalcInfo { get; set; }
        }

        class Provider                                                                              // Клас Доставчик
        {
            public string Name { get; set; }
            public Dictionary<int, Machine> MachinesOld { get; set; }
            public List<Machine> MachinesNew { get; set; }
            public double? CoefOfInflation { get; set; }
            public double? CoefOfAllowedInfl { get; set; }
            public double? CoefOfMaximumInfl { get; set; }
        }

        class Duplicates
        {
            public int Count { get; set; }
            public List<string> Position { get; set; }
        }

        static void Main(string[] args)
        {
            int count = 0;                                                                         // Извличане на популацията доставчици
            Dictionary<string, Provider> providers = new Dictionary<string, Provider>();
            string filePath = "C:\\Users\\PC\\Documents\\Work task\\ReffPriceV2\\ReffPrice\\bin\\Debug\\DataV.xlsx";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            #region Populating Providers

            Console.WriteLine("Start sequence \"Populating provider list\"");
            int rowRange = GetRowsCount(xlRange);                                                   //Обхват
            int monitorDelimeter = rowRange / 10;

            for (int i = 2; i <= rowRange; i++)
            {
                SequenceMonitor(i, monitorDelimeter);

                string providerStr = xlRange.Cells[i, 10].Value2.ToString();

                if (!providers.Keys.Contains(providerStr))
                {
                    Provider provider = new Provider
                    {
                        Name = providerStr,
                        MachinesOld = new Dictionary<int, Machine>(),
                        MachinesNew = new List<Machine>()
                    };
                    providers[providerStr] = provider;
                }

                Machine oldMachine = new Machine
                {
                    RefNum = int.Parse(xlRange.Cells[i, 2].Value2.ToString()),
                    Type = xlRange.Cells[i, 3].Value2.ToString(),
                    Brand = xlRange.Cells[i, 4].Value2.ToString(),
                    Model = xlRange.Cells[i, 5].Value2.ToString(),
                    Clarification = xlRange.Cells[i, 6].Value2.ToString(),
                    OldPrice = double.Parse(xlRange.Cells[i, 7].Value2.ToString())
                };

                oldMachine.OldPosition = xlRange.Cells[i, 1].Value2.ToString();

                if (providers[providerStr].MachinesOld.ContainsKey((int)oldMachine.RefNum))
                {

                    if (xlRange.Cells[i, 11].Value2 == null)
                    {
                        xlRange.Cells[i, 11].Value2 = "Дублиращ се актив с референтен номер;";
                    }
                    else
                    {
                        xlRange.Cells[i, 11].Value2 += " Дублиращ се актив с референтен номер;";
                    }

                    if (providers[providerStr].MachinesOld[(int)oldMachine.RefNum].OldPrice != 0)
                    {
                        providers[providerStr].MachinesOld[(int)oldMachine.RefNum].OldPrice = 0;

                        int temp = int.Parse(providers[providerStr].MachinesOld[(int)oldMachine.RefNum].OldPosition) + 1;

                        if (xlRange.Cells[temp, 11].Value2 == null)
                        {
                            xlRange.Cells[temp, 11].Value2 = "Дублиращ се актив с референтен номер;";
                        }
                        else
                        {
                            xlRange.Cells[temp, 11].Value2 += " Дублиращ се актив с референтен номер;";
                        }
                    }
                }
                else
                {
                    providers[providerStr].MachinesOld[(int)oldMachine.RefNum] = oldMachine;
                }

            }
            Console.WriteLine(" Done!");

            for (int r = 2; r <= rowRange; r++)
            {
                if (xlRange.Cells[r, 11].Value2 != null)
                {
                    count++;
                }
            }

            foreach (string provider in providers.Keys)
            {
                List<int> toDelete = new List<int>();
                foreach (int reff in providers[provider].MachinesOld.Keys)
                {
                    if (providers[provider].MachinesOld[reff].OldPrice == 0)
                    {
                        toDelete.Add(reff);
                    }
                }
                for (int i = 0; i < toDelete.Count; i++)
                {
                    providers[provider].MachinesOld.Remove(toDelete[i]);
                }
            }

            Console.WriteLine($"There are {count} problem data sectors in Refferent data.");


            #endregion
                                                                                            //Извличане на списък с машини от предишен период
            #region Populating Machinery 

            Dictionary<int, Machine> oldMachDict = new Dictionary<int, Machine>();

            Console.WriteLine("Start sequence \"Populating database for machinery\"");
            int countMonitor = 0;
            monitorDelimeter = providers.Count / 10;

            foreach (string provider in providers.Keys)
            {
                countMonitor++;
                SequenceMonitor(countMonitor, monitorDelimeter);
                foreach (int refNum in providers[provider].MachinesOld.Keys)
                {
                    if (providers[provider].MachinesOld[refNum].RefNum != null && !oldMachDict.ContainsKey((int)providers[provider].MachinesOld[refNum].RefNum))
                    {
                        oldMachDict[(int)providers[provider].MachinesOld[refNum].RefNum] = new Machine
                        {
                            RefNum = providers[provider].MachinesOld[refNum].RefNum,
                            Type = providers[provider].MachinesOld[refNum].Type,
                            Brand = providers[provider].MachinesOld[refNum].Brand,
                            Model = providers[provider].MachinesOld[refNum].Model,
                            Clarification = providers[provider].MachinesOld[refNum].Clarification
                        };
                    }
                }
            }

            Dictionary<string, string> brands = new Dictionary<string, string>();                   //Речник за сравняване на марки
            foreach (int key in oldMachDict.Keys)
            {
                string brand = oldMachDict[key].
                           Brand.Replace(" ", string.Empty).
                           Replace("\"", string.Empty).
                           ToLower();
                if (!brands.Keys.Contains(brand))
                {
                    brands[brand] = oldMachDict[key].Brand;
                }
            }

            Dictionary<string, string> types = new Dictionary<string, string>();                    //Речник за сравняване на тип машини
            foreach (int key in oldMachDict.Keys)
            {
                string type = oldMachDict[key].
                        Type.Replace(" ", string.Empty).
                        Replace("\"", string.Empty).
                        ToLower();
                if (!types.Keys.Contains(type))
                {
                    types[type] = oldMachDict[key].Type;
                }
            }

            Dictionary<string, string> providersStr = new Dictionary<string, string>();             //Речник за сравняване на доставчици
            foreach (string key in providers.Keys)
            {
                string provider = key
                        .Replace(" ", string.Empty).
                        Replace("\"", string.Empty).
                        ToLower();
                if (!providersStr.Keys.Contains(
                    provider))
                {
                    providersStr[provider] = key;
                }
            }
            Console.WriteLine(" Done!");

            #endregion

                                                                                                    //Четене на списък с машини от новия период

            #region Reading new Data

            Excel._Worksheet xlWorksheetRead = xlWorkbook.Sheets[2];
            Excel.Range xlRangeRead = xlWorksheetRead.UsedRange;

            Console.WriteLine("Start sequence \"Reading new Data\"");
            rowRange = GetRowsCount(xlRangeRead);                                                             //Обхват
            monitorDelimeter = rowRange / 10;

            for (int r = 2; r <= rowRange; r++)
            {
                string provider = xlRangeRead.Cells[r, 11].Value2.ToString().Trim();

                SequenceMonitor(r, monitorDelimeter);

                if (xlRangeRead.Cells[r, 7].Value2 == null)                                                  //Проверка дали е отпаднал актива
                {
                    continue;
                }

                if (xlRangeRead.Cells[r, 12].Value2 != null && Droped(xlRangeRead.Cells[r, 12].Value2))                                                  //Проверка дали е отпаднал актива
                {
                    continue;
                }

                provider = HardTryContains(providersStr, provider);

                if (!providers.ContainsKey(provider))                                                       //Проверка дали е нов доставчика
                {
                    Provider newProvider = new Provider
                    {
                        Name = provider,
                        MachinesNew = new List<Machine>()
                    };
                    xlRangeRead.Cells[r, 13].Value2 = "Нов доставчик!; ";
                    providers[newProvider.Name] = newProvider;
                }
                else
                {
                    provider = providers[provider].Name;
                }

                Machine newMachine = new Machine { };

                try
                {
                    newMachine.NewPrice = double.Parse(TestNumber(xlRangeRead.Cells[r, 7].Value2.ToString()));
                    if (newMachine.NewPrice==0)
                    {
                        if (xlRangeRead.Cells[r, 13].Value2 == null)
                        {
                            xlRangeRead.Cells[r, 13].Value2 = "Нулева стойност;";
                        }
                        else
                        {
                            xlRangeRead.Cells[r, 13].Value2 += " Нулева стойност;";
                        };
                        continue;
                    }
                }
                catch (Exception)
                {
                    if (xlRangeRead.Cells[r, 13].Value2 == null)
                    {
                        xlRangeRead.Cells[r, 13].Value2 = "Цената не може да бъде разчетена;";
                    }
                    else
                    {
                        xlRangeRead.Cells[r, 13].Value2 += " Цената не може да бъде разчетена;";
                    };
                }

                if (xlRangeRead.Cells[r, 9].Value2 != null)
                {
                    newMachine.DateFrom = GetDate(xlRangeRead.Cells[r, 9].Value2.ToString().Trim());
                }

                if (xlRangeRead.Cells[r, 10].Value2 != null)
                {
                    newMachine.DateTo = GetDate(xlRangeRead.Cells[r, 10].Value2.ToString().Trim());
                }

                if (xlRangeRead.Cells[r, 2].Value2 != null)                                                  //Проверка дали има запис за референтен номер
                {
                    try
                    {
                        newMachine.RefNum = int.Parse(xlRangeRead.Cells[r, 2].Value2.ToString().Trim());
                    }
                    catch (Exception)
                    {

                        if (xlRangeRead.Cells[r, 13].Value2 == null)
                        {
                            xlRangeRead.Cells[r, 13].Value2 = "Референтния номер не може да бъде разчетен;";
                        }
                        else
                        {
                            xlRangeRead.Cells[r, 13].Value2 += " Референтния номер не може да бъде разчетен;";
                        };
                    }

                }

                if (xlRangeRead.Cells[r, 3].Value2 == null)                                               //Проверка дали има запис за тип машини
                {
                    if (xlRangeRead.Cells[r, 13].Value2 == null)
                    {
                        xlRangeRead.Cells[r, 13].Value2 = "Няма тип!;";
                    }
                    else
                    {
                        xlRangeRead.Cells[r, 13].Value2 += " Няма тип!;";
                    };
                }
                else                                                                                     //Проверка дали типът отговаря на условията и печатане на реултат ако не отговаря
                {
                    string typeStr = xlRangeRead.Cells[r, 3].Value2.ToString().Trim();
                    if (!types.Values.Contains(typeStr))
                    {
                        typeStr = HardTryContains(types, typeStr);
                        if (!types.Values.Contains(typeStr))
                        {
                            if (xlRangeRead.Cells[r, 13].Value2 == null)
                            {
                                xlRangeRead.Cells[r, 13].Value2 = "Нов тип!;";
                            }
                            else
                            {
                                xlRangeRead.Cells[r, 13].Value2 += " Нов тип!;";
                            }
                            newMachine.Type = typeStr;
                        }
                        else
                        {
                            newMachine.Type = typeStr;
                        }
                    }
                    else
                    {
                        newMachine.Type = typeStr;
                    }

                }

                if (xlRangeRead.Cells[r, 4].Value2 == null)                                  //Проверка дали марката е в базата данни
                { newMachine.Brand = "НЕ Е ВЪВЕДЕНО"; }
                else
                {
                    string brandStr = xlRangeRead.Cells[r, 4].Value2.ToString().Trim();
                    if (!brands.Values.Contains(brandStr))
                    {
                        brandStr = HardTryContains(brands, brandStr);
                        if (!brands.Values.Contains(brandStr))
                        {
                            if (xlRangeRead.Cells[r, 13].Value2 == null)
                            {
                                xlRangeRead.Cells[r, 13].Value2 = "Нова марка;";
                            }
                            else
                            {
                                xlRangeRead.Cells[r, 13].Value2 += " Нова марка;";
                            }
                            newMachine.Brand = brandStr;
                        }
                        else
                        {
                            newMachine.Brand = brandStr;
                        }
                    }
                    else
                    {
                        newMachine.Brand = brandStr;
                    }
                }

                if (xlRangeRead.Cells[r, 5].Value2 != null)
                {
                    newMachine.Model = xlRangeRead.Cells[r, 5].Value2.ToString().Trim();
                }
                else
                {
                    newMachine.Model = "НЕ Е ВЪВЕДЕНО";
                }

                if (xlRangeRead.Cells[r, 6].Value2 != null)
                {
                    newMachine.Clarification = xlRangeRead.Cells[r, 6].Value2.ToString().Trim();
                }
                else
                {
                    newMachine.Clarification = "НЕ Е ВЪВЕДЕНО";
                }

                if (newMachine.RefNum != null && oldMachDict.Keys.Contains((int)newMachine.RefNum))          //Заготовка дали има некоректност в референтния номер. Взимане на данните от базата данни
                {

                    if (!(newMachine.Type == oldMachDict[(int)newMachine.RefNum].Type &&
                        newMachine.Brand == oldMachDict[(int)newMachine.RefNum].Brand &&
                        newMachine.Model == oldMachDict[(int)newMachine.RefNum].Model &&
                        newMachine.Clarification == oldMachDict[(int)newMachine.RefNum].Clarification))
                    {
                        if (xlRangeRead.Cells[r, 13].Value2 == null)
                        {
                            xlRangeRead.Cells[r, 13].Value2 = "Референтния номер не отговаря на базата данни. При последваща обработка референтния номер не е отчетен;";
                        }
                        else
                        {
                            xlRangeRead.Cells[r, 13].Value2 += " Референтния номер не отговаря на базата данни. При последваща обработка референтния номер не е отчетен;";
                        }
                        newMachine.RefNum = null;
                    }
                }

                if (newMachine.RefNum == null)
                {
                    foreach (int reff in providers[provider].MachinesOld.Keys)
                    {
                        bool isModelEqual = false;
                        bool isEqual = false;
                        if (providers[provider].MachinesOld[reff].Type == newMachine.Type &&
                            providers[provider].MachinesOld[reff].Brand == newMachine.Brand)
                        {
                            string oldModel = providers[provider].MachinesOld[reff].Model.ToLower().Replace(" ", string.Empty).Replace(System.Environment.NewLine, string.Empty).Trim('\n').Trim();
                            string newModel = newMachine.Model.ToLower().Replace(" ", string.Empty).Replace(System.Environment.NewLine, string.Empty).Trim('\n');
                            if (oldModel.Length == newModel.Length)
                            {
                                isModelEqual = true;
                                for (int i = 0; i < oldModel.Length; i++)
                                {
                                    if (oldModel[i] == newModel[i])
                                    {

                                    }
                                    else if (HardSpell(oldModel[i], newModel[i]))
                                    {

                                    }
                                    else
                                    {
                                        isModelEqual = false;
                                        break;
                                    }
                                }
                                if (isModelEqual)
                                {
                                    string oldClarification = providers[provider].MachinesOld[reff].Clarification.ToLower().Replace(" ", string.Empty).Replace(System.Environment.NewLine, string.Empty).Trim('\n').Trim();
                                    string newClarification = newMachine.Clarification.ToLower().Replace(" ", string.Empty).Replace(System.Environment.NewLine, string.Empty).Trim('\n');
                                    if (oldClarification.Length == newClarification.Length)
                                    {
                                        isEqual = true;
                                        for (int k = 0; k < oldClarification.Length; k++)
                                        {
                                            if (oldClarification[k] == newClarification[k])
                                            {

                                            }
                                            else if (HardSpell(oldClarification[k], newClarification[k]))
                                            {

                                            }
                                            else
                                            {
                                                isEqual = false;
                                                break;
                                            }
                                        }
                                        if (isEqual && isModelEqual)
                                        {
                                            newMachine.RefNum = reff;
                                            newMachine.Clarification = providers[provider].MachinesOld[reff].Clarification;
                                            newMachine.Model = providers[provider].MachinesOld[reff].Model;
                                            if (xlRangeRead.Cells[r, 13].Value2 == null)
                                            {
                                                xlRangeRead.Cells[r, 13].Value2 = $"Намерена е машина в базата данни, отговаряща като тип, марка, модел в ред {providers[provider].MachinesOld[reff].OldPosition} и уточнение на тази. Взет е референтния и номер при последваща обработка;";
                                            }
                                            else
                                            {
                                                xlRangeRead.Cells[r, 13].Value2 += $" Намерена е машина в базата данни, отговаряща като тип, марка, модел в ред {providers[provider].MachinesOld[reff].OldPosition} и уточнение на тази. Взет е референтния и номер при последваща обработка;";
                                            }
                                        }

                                    }
                                }

                            }
                        }
                        if (isModelEqual && isEqual)
                        {
                            break;
                        }
                    }
                }

                if (xlRangeRead.Cells[r, 8].Value2 != null)
                {
                    try
                    {
                        newMachine.NominalHP = double.Parse(TestNumber(xlRangeRead.Cells[r, 8].Value2.ToString()));
                        if (newMachine.NominalHP==0)
                        {
                            newMachine.NominalHP = null;
                        }
                    }
                    catch (Exception)
                    {
                        if (xlRangeRead.Cells[r, 13].Value2 == null)
                        {
                            xlRangeRead.Cells[r, 13].Value2 = "Конските сили не могат да бъдат разчетени;";
                        }
                        else
                        {
                            xlRangeRead.Cells[r, 13].Value2 += " Конските сили не могат да бъдат разчетени;";
                        }
                    };
                }

                if (newMachine.RefNum != null && oldMachDict.ContainsKey((int)newMachine.RefNum) &&
                    providers[provider].MachinesOld.Count != 0
                    && providers[provider].MachinesOld.ContainsKey((int)newMachine.RefNum))
                {
                    newMachine.OldPrice = providers[provider].MachinesOld[(int)newMachine.RefNum].OldPrice;
                    newMachine.Inflation = newMachine.NewPrice / newMachine.OldPrice;

                }

                newMachine.NewPosition = xlRangeRead.Cells[r, 1].Value2.ToString();
                providers[provider].MachinesNew.Add(newMachine);

            }

            Dictionary<string, Dictionary<string, Duplicates>> chekDuplicates = new Dictionary<string, Dictionary<string, Duplicates>>();
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                if (!chekDuplicates.Keys.Contains(provider))
                {
                    chekDuplicates[provider] = new Dictionary<string, Duplicates>();
                }
                for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                {
                    string temp = providers[provider].MachinesNew[i].Type + providers[provider].MachinesNew[i].Brand + providers[provider].MachinesNew[i].Model + providers[provider].MachinesNew[i].Clarification;
                    if (!chekDuplicates[provider].Keys.Contains(temp))
                    {
                        Duplicates newDuplicate = new Duplicates
                        {
                            Count = 1,
                            Position = new List<string>()
                        };
                        newDuplicate.Position.Add(providers[provider].MachinesNew[i].NewPosition);
                    }
                    else
                    {
                        chekDuplicates[provider][temp].Count++;
                        chekDuplicates[provider][temp].Position.Add(providers[provider].MachinesNew[i].NewPosition);
                    }
                }
            }

            foreach (string provider in chekDuplicates.Keys)
            {
                foreach (string parameteters in chekDuplicates[provider].Keys)
                {
                    if (chekDuplicates[provider][parameteters].Count > 1)
                    {
                        int[] rows = chekDuplicates[provider][parameteters].Position.Select(x => int.Parse(x)).ToArray();
                        for (int i = 0; i < rows.Length; i++)
                        {
                            if (xlRangeRead.Cells[rows[i] + 1, 13].Value2 == null)
                            {
                                xlRangeRead.Cells[rows[i] + 1, 13].Value2 = "Дублиращ се актив!;";
                            }
                            else
                            {
                                xlRangeRead.Cells[rows[i] + 1, 13].Value2 += " Дублиращ се актив!;";
                            }
                        }
                    }
                }

            }

            Console.WriteLine(" Done!");

            count = 0;

            for (int r = 2; r <= rowRange; r++)
            {
                if (xlRangeRead.Cells[r, 12].Value2 != null)
                {
                    count++;
                }
            }

            Console.WriteLine($"There are {count} problem data sectors in Input data.?");

            List<double> formulaData = new List<double>();
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)  //Изчисляване на средния процент на нарастване на цените за даден търговец
            {
                formulaData = new List<double>();
                for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                {
                    if (providers[provider].MachinesNew[i].Inflation != null && providers[provider].MachinesNew[i].Inflation > 1)
                    {
                        formulaData.Add((double)providers[provider].MachinesNew[i].Inflation);
                    }
                }
                if (formulaData.Count > 0)
                {
                    providers[provider].CoefOfInflation = formulaData.Average();
                }
            }

            #endregion

            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        if (providers[provider].MachinesNew[i].Inflation != null && providers[provider].MachinesNew[i].Inflation > 1)
                        {
                            formulaData.Add((double)providers[provider].MachinesNew[i].Inflation);
                        }
                    }
                }
            }
            double totalInflationIndex = 0;
            if (formulaData.Count != 0)
            {
                totalInflationIndex = formulaData.Average();                                 //Изчисляване на средния процент на нарастване на цените на всички активи, за който имаме данни за старите и новите цени от даден търговец
            }
            double MaxAllowedInfl = 1.5;                                                   //Максимално допустимо нарастване на цените за даден търговец (% Х 100)

            double totalInflationIndexCorrection = 0;                                       //Изчисляване на средния процент на нарастване на цените на всички активи, коригиран съобразно максимално допустимото нарастване на цените

            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        if (providers[provider].MachinesNew[i].Inflation != null && providers[provider].MachinesNew[i].Inflation <= MaxAllowedInfl && providers[provider].MachinesNew[i].Inflation > 1)
                        {
                            formulaData.Add((double)providers[provider].MachinesNew[i].Inflation);
                        }
                    }
                }
            }
            if (formulaData.Count == 0) //za celite na TESTA!!
            {
                totalInflationIndexCorrection = MaxAllowedInfl;
            }
            else
            {
                totalInflationIndexCorrection = formulaData.Average();
            }


            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    formulaData = new List<double>();
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        if (providers[provider].MachinesNew[i].Inflation <= MaxAllowedInfl && providers[provider].MachinesNew[i].Inflation > 1)
                        {
                            formulaData.Add((double)providers[provider].MachinesNew[i].Inflation);
                        }
                    }
                    if (formulaData.Count != 0)
                    {
                        providers[provider].CoefOfAllowedInfl = formulaData.Average();
                    }

                }
            }

            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    count = providers[provider].MachinesNew.Count;
                    for (int i = 0; i < count; i++)
                    {
                        if (providers[provider].MachinesNew[i].NewPrice != null)
                        {
                            formulaData.Add(1);
                        }
                    }
                }
            }
            int newMachinesCount = formulaData.Count;

            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    count = providers[provider].MachinesNew.Count;
                    for (int i = 0; i < count; i++)
                    {
                        if (providers[provider].MachinesNew[i].NewPrice != null && providers[provider].MachinesNew[i].OldPrice != null)
                        {
                            formulaData.Add(1);
                        }
                    }
                }
            }
            int newMachinesReffCount = formulaData.Count;

            double reffCoefOld= Math.Ceiling(((double)newMachinesReffCount / (double)newMachinesCount)*100)/100;
            double reffCoefNew = 1 - reffCoefOld;

            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys) //Редуциране на цената на доставчиците до пределно доспустимото нарастване за активи, за които доставчикът е представил цени и в предишен период
            {
                if (providers[provider].CoefOfInflation != null)
                {
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        if (providers[provider].MachinesNew[i].Inflation != null &&
                            providers[provider].MachinesNew[i].Inflation > (double)totalInflationIndexCorrection)
                        {
                            providers[provider].MachinesNew[i].AlowedPrice = providers[provider].MachinesNew[i].OldPrice * (double)totalInflationIndexCorrection;
                            providers[provider].MachinesNew[i].ReffPrice = reffCoefOld * providers[provider].MachinesNew[i].OldPrice + reffCoefNew * providers[provider].MachinesNew[i].AlowedPrice; //Определяне на референтна цена
                            providers[provider].MachinesNew[i].CalcInfo = $" = {reffCoefOld:0.00}*(стара цена - ) {providers[provider].MachinesNew[i].OldPrice:0.00} + {reffCoefNew:0.00}*(стара цена - ) {providers[provider].MachinesNew[i].OldPrice:0.00}*(Коефициент на позволена инфлация - ) {(double)totalInflationIndexCorrection:0.0000}";
                        }
                        else if ((providers[provider].MachinesNew[i].Inflation != null &&
                            providers[provider].MachinesNew[i].Inflation <= (double)totalInflationIndexCorrection) &&
                            providers[provider].MachinesNew[i].Inflation > 1)
                        {
                            providers[provider].MachinesNew[i].ReffPrice = reffCoefOld * providers[provider].MachinesNew[i].OldPrice + reffCoefNew * providers[provider].MachinesNew[i].NewPrice;
                            providers[provider].MachinesNew[i].CalcInfo = $" = {reffCoefOld:0.00}*(стара цена - ) {providers[provider].MachinesNew[i].OldPrice:0.00} +{reffCoefNew:0.00}*(нова цена - ) {providers[provider].MachinesNew[i].NewPrice:0.00}";
                        }
                        else if ((providers[provider].MachinesNew[i].Inflation != null &&
                            providers[provider].MachinesNew[i].Inflation <= (double)totalInflationIndexCorrection) &&
                            providers[provider].MachinesNew[i].Inflation <= 1)
                        {
                            providers[provider].MachinesNew[i].ReffPrice = providers[provider].MachinesNew[i].NewPrice;
                            providers[provider].MachinesNew[i].CalcInfo = $"Новата цена {providers[provider].MachinesNew[i].NewPrice:0.00} е по ниска или равна на старата референтна цена {providers[provider].MachinesNew[i].OldPrice:0.00}";
                        }
                    }
                }
            }

            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys) //Редуциране на цената на доставчиците до пределно доспустимото нарастване за активи, представени от доставчици с частично обновена листа
            {
                if (providers[provider].CoefOfInflation != null)
                {
                    formulaData = new List<double>();
                    double indexForNewMachine = 0;
                    for (int k = 0; k < providers[provider].MachinesNew.Count; k++)
                    {
                        if (providers[provider].MachinesNew[k].Inflation != null &&
                            providers[provider].MachinesNew[k].ReffPrice != null &&
                            providers[provider].MachinesNew[k].NewPrice > providers[provider].MachinesNew[k].ReffPrice)
                        {
                            formulaData.Add((double)providers[provider].MachinesNew[k].ReffPrice / (double)providers[provider].MachinesNew[k].NewPrice);
                        }
                    }
                    if (formulaData.Count != 0)
                    {
                        indexForNewMachine = formulaData.Average();
                        providers[provider].CoefOfMaximumInfl = indexForNewMachine;
                    }
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        if (providers[provider].MachinesNew[i].Inflation == null)
                        {
                            providers[provider].MachinesNew[i].ReffPrice = providers[provider].MachinesNew[i].NewPrice * indexForNewMachine; //Определяне на референтна цена
                            providers[provider].MachinesNew[i].CalcInfo = $" = (нова цена - ){providers[provider].MachinesNew[i].NewPrice:0.00} * (коефициент за максимално нарастване за даден доставчик - ) {indexForNewMachine:0.0000}";
                        }
                    }
                }
            }

            formulaData = new List<double>();
            double totalReduction = 0;                                                                                                     //Редуциране на цената на доставчиците до пределно доспустимото нарастване за активи, представени от доставчици с изцяло обновена листа
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                if (providers[provider].CoefOfInflation != null)
                {
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        if (providers[provider].MachinesNew[i].ReffPrice != null &&
                            providers[provider].MachinesNew[i].NewPrice > providers[provider].MachinesNew[i].ReffPrice)
                        {
                            formulaData.Add((double)providers[provider].MachinesNew[i].ReffPrice / (double)providers[provider].MachinesNew[i].NewPrice);
                        }
                    }
                }
            }
            if (formulaData.Count > 0)
            {
                totalReduction = formulaData.Average();
            }

            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                if (providers[provider].CoefOfInflation == null)
                {
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        providers[provider].MachinesNew[i].ReffPrice = providers[provider].MachinesNew[i].NewPrice * totalReduction;   //Определяне на референтна цена
                        providers[provider].MachinesNew[i].CalcInfo = $"(нова цена -){providers[provider].MachinesNew[i].NewPrice:0.00}* (коефициент за максимална допустима инфлация за всички доставчици) {totalReduction:0.0000}";
                    }
                }
            }

            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        formulaData.Add((double)providers[provider].MachinesNew[i].ReffPrice / (double)providers[provider].MachinesNew[i].NewPrice);
                    }
                }
            }
            double reffReduction = formulaData.Average();                                                                              //Коефициент на средна редукция на референтните цени спрямо представените през 2017 година цени

            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                    {
                        if (providers[provider].MachinesNew[i].OldPrice!=null && providers[provider].MachinesNew[i].ReffPrice > providers[provider].MachinesNew[i].OldPrice)
                        {
                            formulaData.Add((double)providers[provider].MachinesNew[i].ReffPrice / (double)providers[provider].MachinesNew[i].OldPrice);
                        }
                    }
                }
            }
            double oldReffReduction = formulaData.Average();                                                                           //Коефициент на нарастване на референтните цени през 2018 спрямо тези през 2016 година



            #region Print Detailed Output

            Excel._Worksheet xlWorksheetOutp = xlWorkbook.Sheets[3];
            Excel.Range xlRangeOutp = xlWorksheetOutp.UsedRange;

            Console.WriteLine("Start sequence \"Generating Detailed Refferent list table\"");

            if (providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys.Count() < 10)
            {
                monitorDelimeter = 1;
            }
            else
            {
                monitorDelimeter = providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys.Count() / 10;
            }

            int currentRow = 1;

            string[] headerOutput = new string[] { "ITEM_ID", "Тип", "Марка", "Модел",
                    "Уточнение", "Номинална мощност (к.с.)","Цена на доставчик", "Референтна Цена", "Валидна от",
                    "Валидна до","Доставчик","Одиторска следа" };
            for (int c = 1; c <= headerOutput.Length; c++)
            {
                xlRangeOutp.Cells[currentRow, c].Value2 = headerOutput[c - 1];
            }
            currentRow++;
            countMonitor = 0;
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                countMonitor++;
                SequenceMonitor(countMonitor, monitorDelimeter);
                count = providers[provider].MachinesNew.Count();
                for (int i = 0; i < count; i++)
                {

                    if (providers[provider].MachinesNew[i].RefNum != null && oldMachDict.ContainsKey((int)providers[provider].MachinesNew[i].RefNum))
                    {
                        xlRangeOutp.Cells[currentRow, 1].Value2 = providers[provider].MachinesNew[i].RefNum.ToString();
                    }
                    xlRangeOutp.Cells[currentRow, 2].Value2 = providers[provider].MachinesNew[i].Type;
                    xlRangeOutp.Cells[currentRow, 3].Value2 = providers[provider].MachinesNew[i].Brand;
                    xlRangeOutp.Cells[currentRow, 4].Value2 = providers[provider].MachinesNew[i].Model;
                    xlRangeOutp.Cells[currentRow, 5].Value2 = providers[provider].MachinesNew[i].Clarification;
                    if (providers[provider].MachinesNew[i].NominalHP != null)
                    {
                        xlRangeOutp.Cells[currentRow, 6].Value2 = providers[provider].MachinesNew[i].NominalHP;
                    }
                    xlRangeOutp.Cells[currentRow, 7].Value2 = String.Format("{0:0.00}", providers[provider].MachinesNew[i].NewPrice);
                    xlRangeOutp.Cells[currentRow, 8].Value2 = String.Format("{0:0.00}", providers[provider].MachinesNew[i].ReffPrice);
                    xlRangeOutp.Cells[currentRow, 9].Value2 = providers[provider].MachinesNew[i].DateFrom;
                    xlRangeOutp.Cells[currentRow, 10].Value2 = providers[provider].MachinesNew[i].DateTo;
                    xlRangeOutp.Cells[currentRow, 11].Value2 = provider;
                    xlRangeOutp.Cells[currentRow, 12].Value2 = providers[provider].MachinesNew[i].CalcInfo;
                    currentRow++;
                }
            }
            currentRow++;
            formulaData = new List<double>();
            double reffHPpriceTracktor = 0;
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                {
                    if (providers[provider].MachinesNew[i].Type == "Трактори" && providers[provider].MachinesNew[i].NominalHP != null && providers[provider].MachinesNew[i].ReffPrice != null)
                    {
                        formulaData.Add((double)providers[provider].MachinesNew[i].ReffPrice / (double)providers[provider].MachinesNew[i].NominalHP);
                        count++;
                    }
                }
            }
            if (formulaData.Count == 0)
            {
                reffHPpriceTracktor = 0;
            }
            else
            {
                reffHPpriceTracktor = formulaData.Average();
            }
            xlRangeOutp.Cells[currentRow, 1].Value2 = "Цена за 1 конска сила Трактори:";
            xlRangeOutp.Cells[currentRow, 2].Value2 = String.Format("{0:0.00}", reffHPpriceTracktor);
            currentRow += 2;

            formulaData = new List<double>();
            double reffHPpriceHarrvester = 0;
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                {
                    if (providers[provider].MachinesNew[i].Type == "Комбайни" && providers[provider].MachinesNew[i].NominalHP != null && providers[provider].MachinesNew[i].ReffPrice != null)
                    {
                        formulaData.Add((double)providers[provider].MachinesNew[i].ReffPrice / (double)providers[provider].MachinesNew[i].NominalHP);
                        count++;
                    }
                }
            }
            if (formulaData.Count == 0)
            {
                reffHPpriceHarrvester = 0;
            }
            else
            {
                reffHPpriceHarrvester = formulaData.Average();
            }
            xlRangeOutp.Cells[currentRow, 1].Value2 = "Цена за 1 конска сила Комбайни:";
            xlRangeOutp.Cells[currentRow, 2].Value2 = String.Format("{0:0.00}", reffHPpriceHarrvester);
            currentRow += 2;

            formulaData = new List<double>();
            double reffHPpriceLoader = 0;
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                {
                    if (providers[provider].MachinesNew[i].Type != "Комбайни" && providers[provider].MachinesNew[i].Type != "Трактори" && providers[provider].MachinesNew[i].NominalHP != null && providers[provider].MachinesNew[i].ReffPrice != null)
                    {
                        formulaData.Add((double)providers[provider].MachinesNew[i].ReffPrice / (double)providers[provider].MachinesNew[i].NominalHP);
                        count++;
                    }
                }
            }
            if (formulaData.Count == 0)
            {
                reffHPpriceLoader = 0;
            }
            else
            {
                reffHPpriceLoader = formulaData.Average();
            }
            xlRangeOutp.Cells[currentRow, 1].Value2 = "Цена за 1 конска сила Челни товарачи:";
            xlRangeOutp.Cells[currentRow, 2].Value2 = String.Format("{0:0.00}", reffHPpriceLoader);

            Console.WriteLine(" Done!");

            #endregion

            #region Print Second Analis
            Excel._Worksheet xlWorksheetDoubChek = xlWorkbook.Sheets[4];
            Excel.Range xlRangeDoubChek = xlWorksheetDoubChek.UsedRange;

            Console.WriteLine("Start sequence \"Generating Data for Double Chek\"");

            if (providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys.Count() < 10)
            {
                monitorDelimeter = 1;
            }
            else
            {
                monitorDelimeter = providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys.Count() / 10;
            }

            currentRow = 6;

            countMonitor = 0;

            xlRangeDoubChek.Cells[1, 7].Value2 = String.Format("{0}", (MaxAllowedInfl - 1) * 100);
            xlRangeDoubChek.Cells[1, 11].Value2 = String.Format("{0}", (totalInflationIndex - 1) * 100);
            xlRangeDoubChek.Cells[1, 14].Value2 = String.Format("{0}", (totalInflationIndexCorrection - 1) * 100);
            xlRangeDoubChek.Cells[1, 17].Value2 = String.Format("{0}", (totalReduction - 1) * 100);
            xlRangeDoubChek.Cells[1, 20].Value2 = String.Format("{0}", reffCoefOld * 100);
            xlRangeDoubChek.Cells[1, 23].Value2 = String.Format("{0}", (oldReffReduction-1) * 100);
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                countMonitor++;
                SequenceMonitor(countMonitor, monitorDelimeter);

                xlRangeDoubChek.Cells[currentRow, 10].Value2 = provider;
                xlRangeDoubChek.Cells[currentRow, 11].Value2 = "Доставчик";
                xlRangeDoubChek.Cells[currentRow, 22].Value2 = String.Format("{0}", (providers[provider].CoefOfInflation - 1) * 100);
                xlRangeDoubChek.Cells[currentRow, 24].Value2 = String.Format("{0}", (providers[provider].CoefOfAllowedInfl - 1) * 100);
                xlRangeDoubChek.Cells[currentRow, 26].Value2 = String.Format("{0}", (providers[provider].CoefOfMaximumInfl - 1) * 100);

                int newProviderRow = currentRow;
                currentRow++;
                for (int i = 0; i < providers[provider].MachinesNew.Count; i++)
                {
                    xlRangeDoubChek.Cells[currentRow, 1].Value2 = providers[provider].MachinesNew[i].NewPosition;
                    if (providers[provider].MachinesNew[i].RefNum != null && oldMachDict.ContainsKey((int)providers[provider].MachinesNew[i].RefNum))
                    {
                        xlRangeDoubChek.Cells[currentRow, 2].Value2 = providers[provider].MachinesNew[i].RefNum.ToString();
                    }
                    xlRangeDoubChek.Cells[currentRow, 3].Value2 = providers[provider].MachinesNew[i].Type;
                    xlRangeDoubChek.Cells[currentRow, 4].Value2 = providers[provider].MachinesNew[i].Brand;
                    xlRangeDoubChek.Cells[currentRow, 5].Value2 = providers[provider].MachinesNew[i].Model;
                    xlRangeDoubChek.Cells[currentRow, 6].Value2 = providers[provider].MachinesNew[i].Clarification;
                    if (providers[provider].MachinesNew[i].NominalHP != null)
                    {
                        xlRangeDoubChek.Cells[currentRow, 7].Value2 = providers[provider].MachinesNew[i].NominalHP;
                    }
                    xlRangeDoubChek.Cells[currentRow, 8].Value2 = String.Format("{0}", providers[provider].MachinesNew[i].NewPrice);
                    xlRangeDoubChek.Cells[currentRow, 37].Value2 = String.Format("{0}", providers[provider].MachinesNew[i].ReffPrice);
                    xlRangeDoubChek.Cells[currentRow, 9].Value2 = providers[provider].MachinesNew[i].DateFrom;
                    xlRangeDoubChek.Cells[currentRow, 10].Value2 = providers[provider].MachinesNew[i].DateTo;
                    xlRangeDoubChek.Cells[currentRow, 11].Value2 = provider;

                    if (providers[provider].MachinesNew[i].RefNum != null && providers[provider].MachinesOld.ContainsKey((int)providers[provider].MachinesNew[i].RefNum))
                    {
                        xlRangeDoubChek.Cells[currentRow, 13].Value2 = providers[provider].MachinesOld[(int)providers[provider].MachinesNew[i].RefNum].OldPosition;
                        xlRangeDoubChek.Cells[currentRow, 14].Value2 = providers[provider].MachinesOld[(int)providers[provider].MachinesNew[i].RefNum].RefNum;
                        xlRangeDoubChek.Cells[currentRow, 15].Value2 = providers[provider].MachinesOld[(int)providers[provider].MachinesNew[i].RefNum].Type;
                        xlRangeDoubChek.Cells[currentRow, 16].Value2 = providers[provider].MachinesOld[(int)providers[provider].MachinesNew[i].RefNum].Brand;
                        xlRangeDoubChek.Cells[currentRow, 17].Value2 = providers[provider].MachinesOld[(int)providers[provider].MachinesNew[i].RefNum].Model;
                        xlRangeDoubChek.Cells[currentRow, 18].Value2 = providers[provider].MachinesOld[(int)providers[provider].MachinesNew[i].RefNum].Clarification;
                        xlRangeDoubChek.Cells[currentRow, 19].Value2 = providers[provider].MachinesOld[(int)providers[provider].MachinesNew[i].RefNum].OldPrice;
                    }

                    currentRow++;
                }
                xlRangeDoubChek.Cells[newProviderRow + 1, 23].Value2 = $"=AVERAGE(U{newProviderRow + 1}:U{currentRow - 1})";
                xlRangeDoubChek.Cells[newProviderRow + 1, 25].Value2 = $"=AVERAGEIFS(U{newProviderRow + 1}:U{currentRow - 1},U{newProviderRow + 1}:U{currentRow - 1},\">0\"" +
                    $",U{newProviderRow + 1}:U{currentRow - 1},\"<=50\")";
                xlRangeDoubChek.Cells[newProviderRow + 1, 28].Value2 = $"=AVERAGE(AA{newProviderRow + 1}:AA{currentRow - 1})";
                if (providers[provider].CoefOfInflation!=null)
                {
                    int cellAbs = newProviderRow + 1;
                    for (newProviderRow = newProviderRow + 1; newProviderRow <= currentRow - 1; newProviderRow++)
                    {
                        xlRangeDoubChek.Cells[newProviderRow, 32].Value2 = $"=IF(AND(AD{newProviderRow}=\"\",AE{newProviderRow}=\"\",L{newProviderRow}<>\"Доставчик\"),H{newProviderRow}*(100+$AB${cellAbs})/100,\"\")";
                    }

                }

            }
            Console.WriteLine(" Done!");

            #endregion

            Excel._Worksheet xlWorksheetApp1 = xlWorkbook.Sheets[5];
            Excel.Range xlRangetApp1 = xlWorksheetApp1.UsedRange;

            xlRangetApp1.Cells[1, 1].Value2 = "Списък на търговци на земеделска техника, подали информация за ценовите си листи през 2016 г., на които са изпратени и запитвания за актуални цени на цялата им гама машини, съоръжения и/или специализирани транспортни средства";
            currentRow = 2;
            int cellShift = 1;
            count = 1;
            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                xlRangetApp1.Cells[currentRow, cellShift].Value2 = $"{count}";
                xlRangetApp1.Cells[currentRow, cellShift + 1].Value2 = $"{provider}";
                count++;
                currentRow++;
                if (currentRow - 1 == providers.Keys.Count / 2)
                {
                    cellShift += 2;
                    currentRow = 2;
                    formulaData.Add(1);
                }
            }

            int oldProvidersCount = formulaData.Count();

            Excel._Worksheet xlWorksheetApp2 = xlWorkbook.Sheets[6];
            Excel.Range xlRangetApp2 = xlWorksheetApp2.UsedRange;

            xlRangetApp2.Cells[1, 1].Value2 = "Списък на търговци на земеделска техника, подали информация за актуални цени през 2017г. на цялата им гама машини, съоръжения и/или специализирани транспортни средства";
            currentRow = 2;
            cellShift = 1;
            count = 1;
            formulaData = new List<double>();
            foreach (string provider in providers.Keys)
            {
                if (providers[provider].MachinesNew.Count != 0)
                {
                    xlRangetApp2.Cells[currentRow, cellShift].Value2 = $"{count}";
                    xlRangetApp2.Cells[currentRow, cellShift + 1].Value2 = $"{provider}";
                    count++;
                    currentRow++;
                    formulaData.Add(1);
                    if (currentRow - 1 == providers.Where(x => x.Value.MachinesNew != null).ToDictionary(x => x.Key, x => x.Value).Keys.Count / 2)
                    {
                        cellShift += 2;
                        currentRow = 2;
                    }
                }
            }
            int newProvidersCount = formulaData.Count();

            Excel._Worksheet xlWorksheetAnalis = xlWorkbook.Sheets[7];
            Excel.Range xlRangeAnalis = xlWorksheetAnalis.UsedRange;
            currentRow = 1;
            xlRangeAnalis.Cells[currentRow, 1].Value2 = "Доставчик";
            xlRangeAnalis.Cells[currentRow, 2].Value2 = "Коефициент на нарастване на цените за даден доставчик";
            xlRangeAnalis.Cells[currentRow, 3].Value2 = "Коефициент на нарастване на цените за даден доставчик до 50%";
            xlRangeAnalis.Cells[currentRow, 4].Value2 = "Коефициент на позволено нарастване на цените за даден доставчик с частично обновена продуктова листа";
            currentRow++;
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)       //Принтиране на коефициентите на нарастване на цената за всеки доставчик
            {
                if (providers[provider].CoefOfInflation != null)
                {
                    xlRangeAnalis.Cells[currentRow, 1].Value2 = provider;
                    xlRangeAnalis.Cells[currentRow, 2].Value2 = String.Format("{0:0.0000}", providers[provider].CoefOfInflation);
                    if (providers[provider].CoefOfAllowedInfl != null)
                    {
                        xlRangeAnalis.Cells[currentRow, 3].Value2 = String.Format("{0:0.0000}", providers[provider].CoefOfAllowedInfl);
                    }
                    if (providers[provider].CoefOfMaximumInfl != null)
                    {
                        xlRangeAnalis.Cells[currentRow, 4].Value2 = String.Format("{0:0.0000}", providers[provider].CoefOfMaximumInfl);
                    }
                    currentRow++;
                }
            }
            currentRow++;
            xlRangeAnalis.Cells[currentRow, 1].Value2 = "Коефициент на нарастване на цените за всички активи";
            xlRangeAnalis.Cells[currentRow, 2].Value2 = $"{totalInflationIndex:0.0000}";
            currentRow++;
            xlRangeAnalis.Cells[currentRow, 1].Value2 = $"Коефициент на нарастване на цените за всички активи, без тези на доставчици с коефициент по - голям от {MaxAllowedInfl:0.0000}";
            xlRangeAnalis.Cells[currentRow, 2].Value2 = $"{totalInflationIndexCorrection:0.0000}";
            currentRow++;
            xlRangeAnalis.Cells[currentRow, 1].Value2 = $"Коефициент на нарастване на цените за всички активи за доставчици с изцяло обновена продуктова листа";
            xlRangeAnalis.Cells[currentRow, 2].Value2 = $"{totalReduction:0.0000}";
            currentRow++;
            xlRangeAnalis.Cells[currentRow, 1].Value2 = $"Коефициент на редукция след прилагане на методиката";
            xlRangeAnalis.Cells[currentRow, 2].Value2 = $"{reffReduction:0.0000}";
            currentRow++;
            xlRangeAnalis.Cells[currentRow, 1].Value2 = $"Коефициент на нарастване на референтните цени";
            xlRangeAnalis.Cells[currentRow, 2].Value2 = $"{oldReffReduction:0.0000}";
            currentRow++;

            #region Print Output

            Excel._Worksheet xlWorksheetOutput = xlWorkbook.Sheets[8];
            Excel.Range xlRangeOutput = xlWorksheetOutput.UsedRange;

            Console.WriteLine("Start sequence \"Generating short Refferent list table\"");

            if (providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys.Count() < 10)
            {
                monitorDelimeter = 1;
            }
            else
            {
                monitorDelimeter = providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys.Count() / 10;
            }

           currentRow = 1;

            headerOutput = new string[] { "Тип", "Марка", "Модел",
                    "Уточнение", "Референтна Цена", "Валидна от",
                    "Валидна до","Доставчик", };
            for (int c = 1; c <= headerOutput.Length; c++)
            {
                xlRangeOutput.Cells[currentRow, c].Value2 = headerOutput[c - 1];
            }
            currentRow++;
            countMonitor = 0;
            foreach (string provider in providers.Where(x => x.Value.MachinesNew.Count != 0).ToDictionary(x => x.Key, x => x.Value).Keys)
            {
                countMonitor++;
                SequenceMonitor(countMonitor, monitorDelimeter);
                count = providers[provider].MachinesNew.Count();
                for (int i = 0; i < count; i++)
                {
                    xlRangeOutput.Cells[currentRow, 1].Value2 = providers[provider].MachinesNew[i].Type;
                    xlRangeOutput.Cells[currentRow, 2].Value2 = providers[provider].MachinesNew[i].Brand;
                    xlRangeOutput.Cells[currentRow, 3].Value2 = providers[provider].MachinesNew[i].Model;
                    xlRangeOutput.Cells[currentRow, 4].Value2 = providers[provider].MachinesNew[i].Clarification;
                    xlRangeOutput.Cells[currentRow, 5].Value2 = String.Format("{0:0.00}", providers[provider].MachinesNew[i].ReffPrice);

                    xlRangeOutput.Cells[currentRow, 8].Value2 = provider;
                    currentRow++;
                }
            }
            Console.WriteLine(" Done!");
            #endregion


            Excel._Worksheet xlWorksheetTracktor = xlWorkbook.Sheets[9];
            Excel.Range xlRangeTracktor = xlWorksheetTracktor.UsedRange;

            Console.WriteLine("Start sequence \"Generating price per HP for Tracktors\"");

            rowRange = GetRowsCount(xlRangeTracktor);

            monitorDelimeter = rowRange / 10;

            currentRow = 2;
            for (int i = 2; i <= rowRange; i++)
            {
                SequenceMonitor(i, monitorDelimeter);

                string providerStr = xlRangeTracktor.Cells[i, 12].Value2.ToString();
                string type = xlRangeTracktor.Cells[i, 3].Value2.ToString();
                string brand = xlRangeTracktor.Cells[i, 4].Value2.ToString();
                string model = xlRangeTracktor.Cells[i, 5].Value2.ToString();
                string clarification = xlRangeTracktor.Cells[i, 6].Value2.ToString();
                for (int k = 0; k < providers[providerStr].MachinesNew.Count; k++)
                {
                    if (providers[providerStr].MachinesNew[k].Type == type &&
                        providers[providerStr].MachinesNew[k].Brand == brand &&
                        providers[providerStr].MachinesNew[k].Model == model &&
                        providers[providerStr].MachinesNew[k].Clarification == clarification)
                    {
                        xlRangeTracktor.Cells[i, 9].Value2 = Math.Round((double)providers[providerStr].MachinesNew[k].ReffPrice,2);
                        xlRangeTracktor.Cells[i, 13].Value2 = providers[providerStr].MachinesNew[k].CalcInfo;
                        break;
                    }
                }
            }
            Console.WriteLine(" Done!");

            Excel._Worksheet xlWorksheetHarvester = xlWorkbook.Sheets[10];
            Excel.Range xlRangeHarvester = xlWorksheetHarvester.UsedRange;

            Console.WriteLine("Start sequence \"Generating price per HP for Harvesters\"");

            rowRange = GetRowsCount(xlRangeHarvester);

            monitorDelimeter = rowRange / 10;

            currentRow = 2;
            for (int i = 2; i <= rowRange; i++)
            {
                SequenceMonitor(i, monitorDelimeter);

                string providerStr = xlRangeHarvester.Cells[i, 12].Value2.ToString();
                string type = xlRangeHarvester.Cells[i, 3].Value2.ToString();
                string brand = xlRangeHarvester.Cells[i, 4].Value2.ToString();
                string model = xlRangeHarvester.Cells[i, 5].Value2.ToString();
                string clarification = xlRangeHarvester.Cells[i, 6].Value2.ToString();
                for (int k = 0; k < providers[providerStr].MachinesNew.Count; k++)
                {
                    if (providers[providerStr].MachinesNew[k].Type == type &&
                        providers[providerStr].MachinesNew[k].Brand == brand &&
                        providers[providerStr].MachinesNew[k].Model == model &&
                        providers[providerStr].MachinesNew[k].Clarification == clarification)
                    {
                        xlRangeHarvester.Cells[i, 9].Value2 = Math.Round((double)providers[providerStr].MachinesNew[k].ReffPrice, 2);
                        xlRangeHarvester.Cells[i, 13].Value2 = providers[providerStr].MachinesNew[k].CalcInfo;
                        break;
                    }
                }
            }
            Console.WriteLine(" Done!");

            Excel._Worksheet xlWorksheetSilHarvester = xlWorkbook.Sheets[11];
            Excel.Range xlRangeSilHarvester = xlWorksheetSilHarvester.UsedRange;

            Console.WriteLine("Start sequence \"Generating price per HP for Silage Harvesters\"");

            rowRange = GetRowsCount(xlRangeSilHarvester);

            monitorDelimeter = rowRange / 10;

            currentRow = 2;
            for (int i = 2; i <= rowRange; i++)
            {
                SequenceMonitor(i, monitorDelimeter);

                string providerStr = xlRangeSilHarvester.Cells[i, 12].Value2.ToString();
                string type = xlRangeSilHarvester.Cells[i, 3].Value2.ToString();
                string brand = xlRangeSilHarvester.Cells[i, 4].Value2.ToString();
                string model = xlRangeSilHarvester.Cells[i, 5].Value2.ToString();
                string clarification = xlRangeSilHarvester.Cells[i, 6].Value2.ToString();
                for (int k = 0; k < providers[providerStr].MachinesNew.Count; k++)
                {
                    if (providers[providerStr].MachinesNew[k].Type == type &&
                        providers[providerStr].MachinesNew[k].Brand == brand &&
                        providers[providerStr].MachinesNew[k].Model == model &&
                        providers[providerStr].MachinesNew[k].Clarification == clarification)
                    {
                        xlRangeSilHarvester.Cells[i, 9].Value2 = Math.Round((double)providers[providerStr].MachinesNew[k].ReffPrice, 2);
                        xlRangeSilHarvester.Cells[i, 13].Value2 = providers[providerStr].MachinesNew[k].CalcInfo;
                        break;
                    }
                }
            }
            Console.WriteLine(" Done!");

            Excel._Worksheet xlWorksheetLoader = xlWorkbook.Sheets[12];
            Excel.Range xlRangeLoader = xlWorksheetLoader.UsedRange;

            Console.WriteLine("Start sequence \"Generating price per HP for Silage Harvesters\"");

            rowRange = GetRowsCount(xlRangeLoader);

            monitorDelimeter = rowRange / 10;

            currentRow = 2;
            for (int i = 2; i <= rowRange; i++)
            {
                SequenceMonitor(i, monitorDelimeter);

                string providerStr = xlRangeLoader.Cells[i, 12].Value2.ToString();
                string type = xlRangeLoader.Cells[i, 3].Value2.ToString();
                string brand = xlRangeLoader.Cells[i, 4].Value2.ToString();
                string model = xlRangeLoader.Cells[i, 5].Value2.ToString();
                string clarification = xlRangeLoader.Cells[i, 6].Value2.ToString();
                for (int k = 0; k < providers[providerStr].MachinesNew.Count; k++)
                {
                    if (providers[providerStr].MachinesNew[k].Type == type &&
                        providers[providerStr].MachinesNew[k].Brand == brand &&
                        providers[providerStr].MachinesNew[k].Model == model &&
                        providers[providerStr].MachinesNew[k].Clarification == clarification)
                    {
                        xlRangeLoader.Cells[i, 9].Value2 = Math.Round((double)providers[providerStr].MachinesNew[k].ReffPrice, 2);
                        xlRangeLoader.Cells[i, 13].Value2 = providers[providerStr].MachinesNew[k].CalcInfo;
                        break;
                    }
                }
            }
            Console.WriteLine(" Done!");

            Excel._Worksheet xlWorksheetSprinckles = xlWorkbook.Sheets[13];
            Excel.Range xlRangeSprinckles = xlWorksheetSprinckles.UsedRange;

            Console.WriteLine("Start sequence \"Generating price per HP for Silage Harvesters\"");

            rowRange = GetRowsCount(xlRangeSprinckles);

            monitorDelimeter = rowRange / 10;

            currentRow = 2;
            for (int i = 2; i <= rowRange; i++)
            {
                SequenceMonitor(i, monitorDelimeter);

                string providerStr = xlRangeSprinckles.Cells[i, 12].Value2.ToString();
                string type = xlRangeSprinckles.Cells[i, 3].Value2.ToString();
                string brand = xlRangeSprinckles.Cells[i, 4].Value2.ToString();
                string model = xlRangeSprinckles.Cells[i, 5].Value2.ToString();
                string clarification = xlRangeSprinckles.Cells[i, 6].Value2.ToString();
                for (int k = 0; k < providers[providerStr].MachinesNew.Count; k++)
                {
                    if (providers[providerStr].MachinesNew[k].Type == type &&
                        providers[providerStr].MachinesNew[k].Brand == brand &&
                        providers[providerStr].MachinesNew[k].Model == model &&
                        providers[providerStr].MachinesNew[k].Clarification == clarification)
                    {
                        xlRangeSprinckles.Cells[i, 9].Value2 = Math.Round((double)providers[providerStr].MachinesNew[k].ReffPrice, 2);
                        xlRangeSprinckles.Cells[i, 13].Value2 = providers[providerStr].MachinesNew[k].CalcInfo;
                        break;
                    }
                }
            }
            Console.WriteLine(" Done!");

            Excel._Worksheet xlWorksheetOthers = xlWorkbook.Sheets[14];
            Excel.Range xlRangeOthers = xlWorksheetOthers.UsedRange;

            Console.WriteLine("Start sequence \"Generating price per HP for Silage Harvesters\"");

            rowRange = GetRowsCount(xlRangeOthers);

            monitorDelimeter = rowRange / 10;

            currentRow = 2;
            for (int i = 2; i <= rowRange; i++)
            {
                SequenceMonitor(i, monitorDelimeter);

                string providerStr = xlRangeOthers.Cells[i, 12].Value2.ToString();
                string type = xlRangeOthers.Cells[i, 3].Value2.ToString();
                string brand = xlRangeOthers.Cells[i, 4].Value2.ToString();
                string model = xlRangeOthers.Cells[i, 5].Value2.ToString();
                string clarification = xlRangeOthers.Cells[i, 6].Value2.ToString();
                for (int k = 0; k < providers[providerStr].MachinesNew.Count; k++)
                {
                    if (providers[providerStr].MachinesNew[k].Type == type &&
                        providers[providerStr].MachinesNew[k].Brand == brand &&
                        providers[providerStr].MachinesNew[k].Model == model &&
                        providers[providerStr].MachinesNew[k].Clarification == clarification)
                    {
                        xlRangeOthers.Cells[i, 9].Value2 = Math.Round((double)providers[providerStr].MachinesNew[k].ReffPrice, 2);
                        xlRangeOthers.Cells[i, 13].Value2 = providers[providerStr].MachinesNew[k].CalcInfo;
                        break;
                    }
                }
            }
            Console.WriteLine(" Done!");

            #region Methodics

            Console.Write("Generating Methodics");

            string methodics = "CLASSIFIED";

            string file_name = @"C:\Users\PC\Documents\Work task\ReffPriceV2\ReffPrice\bin\Debug\Meth.txt";
            System.IO.StreamWriter objWriter;
            objWriter = new System.IO.StreamWriter(file_name);

            objWriter.Write(methodics);
            objWriter.Close();
            Console.WriteLine(" Done!");

#endregion

            xlWorkbook.Save();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Finished!");
            Console.ReadLine();

        }

        private static bool Droped(dynamic value2)
        {
            string methString = value2.ToLower().Trim();
            bool dropped = true;
            string test = "отпад";
            int len = Math.Min(test.Length, methString.Length);
            if (len == 0 || len < test.Length)
            {
                dropped = false;
                return dropped;
            }
            else
            {
                for (int i = 0; i < len; i++)
                {
                    if (methString[i] != test[i])
                    {
                        dropped = false;
                        break;
                    }
                }
            }
            return dropped;
        }

        private static string GetDate(dynamic dynamic)
        {
            string methString = dynamic;
            methString = methString.Replace(',', '.').Trim();
            return methString;
        }

        private static string HardTryContains(Dictionary<string, string> HardDictionary, string tryString)
        {
            string methString = tryString;
            bool isEqual = true;
            if (!HardDictionary.Values.Contains(methString))
            {
                isEqual = false;
                methString = methString.ToLower().Replace(" ", string.Empty).Replace("\"", string.Empty);
            }
            else
            {
                return methString;
            }

            Dictionary<string, string> methDictionary = HardDictionary.
                Where(x => x.Key.Length == methString.Length).
                ToDictionary(x => x.Key, x => x.Value);

            foreach (string key in methDictionary.Keys)
            {
                isEqual = true;
                for (int i = 0; i < key.Length; i++)
                {
                    if (key[i] == methString[i])
                    {

                    }
                    else if (HardSpell(key[i], methString[i]))
                    {

                    }
                    else
                    {
                        isEqual = false;
                        break;
                    }
                }

                if (isEqual == true)
                {
                    methString = HardDictionary[key];
                    return methString;
                }
            }

            if (isEqual == false)
            {
                methString = tryString;
            }

            return methString;
        }

        private static bool HardSpell(char key, char v2)
        {
            bool isEqual = false;
            if ((int)key >= 97 && (int)key <= 122)
            {
                switch (v2)
                {
                    case 'а': v2 = 'a'; break;
                    case 'в': v2 = 'b'; break;
                    case 'с': v2 = 'c'; break;
                    case 'е': v2 = 'e'; break;
                    case 'н': v2 = 'h'; break;
                    case 'к': v2 = 'k'; break;
                    case 'м': v2 = 'm'; break;
                    case 'о': v2 = 'o'; break;
                    case 'р': v2 = 'p'; break;
                    case 'т': v2 = 't'; break;
                    case 'х': v2 = 'x'; break;
                    case 'у': v2 = 'y'; break;
                    default:
                        break;
                }
                if (key == v2)
                {
                    isEqual = true;
                }
            }
            else if ((int)key >= 1072 && (int)key <= 1103)
            {
                switch (v2)
                {
                    case 'a': v2 = 'а'; break;
                    case 'b': v2 = 'в'; break;
                    case 'e': v2 = 'е'; break;
                    case 'k': v2 = 'к'; break;
                    case 'm': v2 = 'м'; break;
                    case 'h': v2 = 'н'; break;
                    case 'o': v2 = 'о'; break;
                    case 'p': v2 = 'р'; break;
                    case 'c': v2 = 'с'; break;
                    case 't': v2 = 'т'; break;
                    case 'x': v2 = 'х'; break;
                    case 'y': v2 = 'у'; break;
                    default:
                        break;
                }
                if (key == v2)
                {
                    isEqual = true;
                }
            }
            return isEqual;
        }

        private static int GetRowsCount(Excel.Range xlRange)
        {
            int count = 1;
            while (xlRange.Cells[count, 1].Value2 != null)
            {
                count++;
            }
            return count - 1;
        }

        private static string TestDouble(string priceStr)
        {
            string output = string.Empty;
            if (priceStr == string.Empty)
            {
                output = "0";
            }
            else
            {
                output = priceStr.Trim().Replace(',', '.').Replace(" ", string.Empty);
            }
            return output;
        }

        private static string TestNumber(string priceStr)
        {
            string output = string.Empty;
            char[] digits = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.' };
            if (priceStr == string.Empty)
            {
                output = "0";
            }
            else
            {
                output = priceStr.Trim().Replace(',', '.').Replace(" ", string.Empty);
                for (int i = 0; i < output.Length; i++)
                {
                    if (!digits.Contains(output[i]))
                    {
                        output = output.Remove(i, 1);
                    }
                }
                while (output[output.Length - 1] == '.')
                {
                    output = output.Remove(output.Length - 1, 1);
                }
            }
            return output;
        }

        private static void SequenceMonitor(int i, int rowRangeDelimeter)
        {
            if (rowRangeDelimeter == 0)
            {
                rowRangeDelimeter = 1;
            }
            if (i % rowRangeDelimeter == 0)
            {
                Console.Write(">");
            }
        }
    }
}
