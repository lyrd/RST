using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RST
{
    //Считывание разметки из файлов и обработка

        //TODO: Удление пробелов
    class WorkWithMarkingFile
    {
        /* TODO: пути к файлам в ini
         * Проверка при чтении строки на корректность
        */
        private static readonly string heatSupplyMarkings = Directory.GetCurrentDirectory() + @"\Markings\heatSupplyMarking.txt";
        private static readonly string waterSupplyMarkings = Directory.GetCurrentDirectory() + @"\Markings\waterSupplyMarking.txt";
        private static readonly string sewerageMarkings = Directory.GetCurrentDirectory() + @"\Markings\sewerageMarking.txt";

        //BD132:К:K115/1000*K1*K3:Смета ТС/const*coeffK1*coeffK3
        //private string targetCell = String.Empty;  //BD132
        //private string targetSheet = String.Empty; //К
        //private string sourceCell = String.Empty;  //K115/1000*K1*K3
        //private string sourceSheet = String.Empty; //Смета ТС/const*coeffK1*coeffK3  

        private List<Marking> listOfMarking = new List<Marking>();

        private string pathToMarkingFile = String.Empty;
        private string pathToCoefficientFile = String.Empty;

        public WorkWithMarkingFile()
        {

        }

        private enum TypeOfExpression
        {
            FORMULA = 1,
            LIST_OF_SOURCES = 2
        }

        private string ConvertExpressionWithPercent(string expressionForConvert, TypeOfExpression type)
        {
            string pattern = String.Empty;
            string replacement = String.Empty;

            pattern = @"([A-zА-я0-9\s]{0,})([/*+-])([A-zА-я0-9\s]{0,})([%])";

            if (type == TypeOfExpression.FORMULA)
            {
                //pattern = @"([A-Z]{0,}[0-9]{1,})([*+-/])([A-Z]{0,}[0-9]{1,})([%])";
                replacement = "$1$2($1*$3/100)";
            }
            else if (type == TypeOfExpression.LIST_OF_SOURCES)
            {
                //pattern = @"([A-zА-я0-9]{1,})([*+-/])([A-zА-я0-9\s]{1,})([%])";
                replacement = "$1$2($1*$3/const)";
            }

            return expressionForConvert = System.Text.RegularExpressions.Regex.Replace(expressionForConvert, pattern, replacement);
        }

        public List<Marking> GetListOfMarkings()
        {
            string[] arrayOfMarking = File.ReadAllLines(Directory.GetCurrentDirectory() + @"\test\test.txt");

            string[] stringSeparators = new string[] { ":" };
            string[] cellsRangeSeparators = new string[] { "@" };
            const char commentTag = '#';

            Console.WriteLine($"***{System.Reflection.MethodBase.GetCurrentMethod().Name}_TEST***");

            foreach (string elem in arrayOfMarking)
            {
                Console.WriteLine($"{elem} is {Array.IndexOf(arrayOfMarking, elem)}");
            }

            foreach (string elem in arrayOfMarking)
            {
                string[] partsOfMarking = elem.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

                string[] targetCellSplited;
                string[] sourceCellSplited;

                if (elem[0] == commentTag) //начало строки # - комментарий
                    continue;
                else if (!partsOfMarking[0].Contains(cellsRangeSeparators[0]) && !partsOfMarking[2].Contains(cellsRangeSeparators[0]))
                {
                    listOfMarking.Add(new Marking(partsOfMarking[0], partsOfMarking[1], partsOfMarking[2], partsOfMarking[3]));
                }
                else if (partsOfMarking[0].Contains(cellsRangeSeparators[0]) && !partsOfMarking[2].Contains(cellsRangeSeparators[0]))
                {
                    targetCellSplited = partsOfMarking[0].Split(cellsRangeSeparators, StringSplitOptions.RemoveEmptyEntries);

                    ExcelUtility.SeparatedCellName firstCellInRange = ExcelUtility.GetSeparatedCellName(targetCellSplited[0]);
                    ExcelUtility.SeparatedCellName secondCellInRange = ExcelUtility.GetSeparatedCellName(targetCellSplited[1]);

                    Console.WriteLine($"Range: {targetCellSplited[0]}\t{targetCellSplited[1]}\r\n{firstCellInRange.Letter}\t{firstCellInRange.Number}\r\n{secondCellInRange.Letter}\t{secondCellInRange.Number}");
                    if (firstCellInRange.Letter == secondCellInRange.Letter)
                    {
                        if (firstCellInRange.Number < secondCellInRange.Number)
                        {
                            for (int i = firstCellInRange.Number; i <= secondCellInRange.Number; i++)
                            {
                                listOfMarking.Add(new Marking($"{firstCellInRange.Letter}{i}", partsOfMarking[1], partsOfMarking[2], partsOfMarking[3]));
                            }
                        }
                        else
                        {
                            //Сообщение в лог (неверный диапазон, начальное значение больше конечного), continue
                            SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{System.Reflection.MethodBase.GetCurrentMethod().Name}\r\nОшибка в формуле \"{elem}\": Неверный диапазон, начальное значение больше конечного \"{targetCellSplited[0]}\" меньше \"{targetCellSplited[1]}\". Формула пропущена\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                            continue;
                        }
                    }
                    else
                    {
                        //Сообщение в лог (неверный диапазон, должен быть один столбец), continue
                        SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{System.Reflection.MethodBase.GetCurrentMethod().Name}\r\nОшибка в формуле \"{elem}\": Неверный диапазон, разные столбцы \"{targetCellSplited[0]}\" != \"{targetCellSplited[1]}\". Формула пропущена\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                        continue;
                    }
                }
                else if (partsOfMarking[0].Contains(cellsRangeSeparators[0]) && partsOfMarking[2].Contains(cellsRangeSeparators[0]))
                {
                    targetCellSplited = partsOfMarking[0].Split(cellsRangeSeparators, StringSplitOptions.RemoveEmptyEntries);
                    sourceCellSplited = partsOfMarking[2].Split(cellsRangeSeparators, StringSplitOptions.RemoveEmptyEntries);

                    ExcelUtility.SeparatedCellName firstTargetCellInRange = ExcelUtility.GetSeparatedCellName(targetCellSplited[0]);
                    ExcelUtility.SeparatedCellName secondTargetCellInRange = ExcelUtility.GetSeparatedCellName(targetCellSplited[1]);

                    ExcelUtility.SeparatedCellName firstSourceCellInRange = ExcelUtility.GetSeparatedCellName(sourceCellSplited[0]);
                    ExcelUtility.SeparatedCellName secondSourceCellInRange = ExcelUtility.GetSeparatedCellName(sourceCellSplited[1]);

                    Console.WriteLine($"Range: {targetCellSplited[0]}\t{targetCellSplited[1]}\r\n{firstTargetCellInRange.Letter}\t{firstTargetCellInRange.Number}\r\n{secondTargetCellInRange.Letter}\t{secondTargetCellInRange.Number}");
                    if (firstTargetCellInRange.Letter == secondTargetCellInRange.Letter)
                    {
                        if (firstTargetCellInRange.Number < secondTargetCellInRange.Number)
                        {
                            if (firstSourceCellInRange.Letter == secondSourceCellInRange.Letter)
                            {
                                if(firstSourceCellInRange.Number < secondSourceCellInRange.Number)
                                {
                                    int index = firstSourceCellInRange.Number;

                                    int deltaTargetCell = secondTargetCellInRange.Number - firstTargetCellInRange.Number;
                                    int deltaSourceCell = secondSourceCellInRange.Number - firstSourceCellInRange.Number;

                                    if (deltaTargetCell == deltaSourceCell)
                                    {
                                        for (int i = firstTargetCellInRange.Number; i <= secondTargetCellInRange.Number; i++)
                                        {
                                            listOfMarking.Add(new Marking($"{firstTargetCellInRange.Letter}{i}", partsOfMarking[1], $"{firstSourceCellInRange.Letter}{index}", partsOfMarking[3]));
                                            index++;
                                        }
                                    }
                                    else
                                    {
                                        SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{System.Reflection.MethodBase.GetCurrentMethod().Name}\r\nОшибка в формуле \"{elem}\": Разные размеры диапазонов. Формула пропущена\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                                        continue;
                                    }
                                }
                                else
                                {
                                    //Сообщение в лог (неверный диапазон, начальное значение больше конечного), continue
                                    SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{System.Reflection.MethodBase.GetCurrentMethod().Name}\r\nОшибка в формуле \"{elem}\": Неверный диапазон, начальное значение больше конечного \"{sourceCellSplited[0]}\" меньше \"{sourceCellSplited[1]}\". Формула пропущена\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                                    continue;
                                }
                            }
                            else
                            {
                                //Сообщение в лог (неверный диапазон, должен быть один столбец), continue
                                SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{System.Reflection.MethodBase.GetCurrentMethod().Name}\r\nОшибка в формуле \"{elem}\": Неверный диапазон, разные столбцы \"{sourceCellSplited[0]}\" != \"{sourceCellSplited[1]}\". Формула пропущена\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                            }
                        }
                        else
                        {
                            //Сообщение в лог (неверный диапазон, начальное значение больше конечного), continue
                            SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{System.Reflection.MethodBase.GetCurrentMethod().Name}\r\nОшибка в формуле \"{elem}\": Неверный диапазон, начальное значение больше конечного \"{targetCellSplited[0]}\" меньше \"{targetCellSplited[1]}\". Формула пропущена\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                            continue;
                        }
                    }
                    else
                    {
                        //Сообщение в лог (неверный диапазон, должен быть один столбец), continue
                        SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{System.Reflection.MethodBase.GetCurrentMethod().Name}\r\nОшибка в формуле \"{elem}\": Неверный диапазон, разные столбцы \"{targetCellSplited[0]}\" != \"{targetCellSplited[1]}\". Формула пропущена\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                        continue;
                    }
                }
            }

            for (int i = 0; i < listOfMarking.Count; i++)
            {
                if (listOfMarking[i].SourceCell.Contains("%"))
                    listOfMarking[i].SourceCell = ConvertExpressionWithPercent(listOfMarking[i].SourceCell, TypeOfExpression.FORMULA);

                if (listOfMarking[i].SourceSheet.Contains("%"))
                    listOfMarking[i].SourceSheet = ConvertExpressionWithPercent(listOfMarking[i].SourceSheet, TypeOfExpression.LIST_OF_SOURCES);
            }

            Console.WriteLine("***listOfMarking***");
            for (int i = 0; i < listOfMarking.Count; i++)
            {
                Console.WriteLine($"{listOfMarking[i].TargetCell}\t{listOfMarking[i].TargetSheet}\t{listOfMarking[i].SourceCell}\t{listOfMarking[i].SourceSheet}");
            }

            return listOfMarking;
        }
    }
}
