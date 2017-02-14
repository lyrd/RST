using MathExpressionsParser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RST
{
    /*
     * TODO: Логирование ошибок
     */ 
    class Program
    {
        public string ClassName
        {
            get
            {
                return GetType().Name;
            }
        }

        readonly private static string extensionOfErrorLogFile = "txt";
        readonly private static string nameOfErrorLogFile = "error_log";

        private static SavesIntoFile savesIntoFile = new SavesIntoFile();

        private struct SeparatedCellNameStruct
        {
            private string letter;
            private short number;            

            public string Letter
            {
                get { return this.letter; }
                set { this.letter = value; }
            }

            public short Number
            {
                get { return this.number; }
                set { this.number = value; }
            }        
        }

        private static string GetFormulaWithValues(StringBuilder data, string[] oldInString, double[] valuesForReplace)
        {
            for (byte index = 0; index < valuesForReplace.Length; index++)
            {
                data.Replace(oldInString[index], valuesForReplace[index].ToString());
            }

            return data.ToString();
        }

        private static double ConvertDecimalSeparator(string s)
        {
            return Double.Parse(s.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
        }

        private enum TypeOfFormula
        {
            FORMULA_WITH_CONST = 1,
            FORMULA_WITH_CONST_CELL = 2,
            FORMULA_WITH_CONST_COEFFICIENT = 3,
            FORMULA_WITH_CONST_CELL_COEFFICIENT = 4,
            FORMULA_WITH_CELL = 5,
            FORMULA_WITH_CELL_COEFFICIENT = 6,
            FORMULA_WITH_COEFFICIENT = 7
        }

        private enum TypeOfVariable
        {
            CELL = 1,
            CONST = 2,
            COEFFICIENT = 3,
            UNKNOW = -1
        }

        private enum ErrorCode
        {
            UNKNOW_VARIABLE_TYPE = 1302201701,
            TRY_PARSE_FAILURE = 1302201702
        }

        private static TypeOfVariable[] GetTypesOfVariables(string[] variablesInFormula)
        {
            TypeOfVariable[] flag = new TypeOfVariable[variablesInFormula.Length];

            for (int i = 0; i < variablesInFormula.Length; i++)
            {
                if (ExcelUtility.CheckData.IsCell(variablesInFormula[i]))
                    flag[i] = TypeOfVariable.CELL;         
                else if (ExcelUtility.CheckData.IsCoefficients(variablesInFormula[i]))
                    flag[i] = TypeOfVariable.COEFFICIENT;
                else if (ExcelUtility.CheckData.IsNumber(variablesInFormula[i]))
                    flag[i] = TypeOfVariable.CONST;
                else
                    flag[i] = TypeOfVariable.UNKNOW;
            }

            return flag;
        }

        private static string MakeCalculation(string formula)
        {
            //Получение имени текущего метода для вывода в логи
            string currMethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

            MathParser mathParser = new MathParser();
            bool isRadians = false;

            //Сепараторы для парсинга формулы
            string[] stringSeparators = new string[] { "+", "-", "/", "*", "(", ")" };

            //Временный массив для хранения распарсенной формулы
            string[] variablesTemp = formula.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

            //Получение типов переменных формулы
            TypeOfVariable[] typeOfVariables = GetTypesOfVariables(variablesTemp);

            //Массив для значений переменных
            double[] values = new double[variablesTemp.Length];

            /*В соответствии с типом переменной для неё выполняются соответсвующе действия:
             * для TypeOfVariable.CELL значение считывается из соответствующей ячейки
             * для TypeOfVariable.COEFFICIENT подставляется расчитанное знчение коэффициента
             * для TypeOfVariable.CONST значение берется из массива переменных, которое передеется как аргумент в этот метод
             * тип TypeOfVariable.UNKNOW присваивается всем переменным, которые не попали ни под один другой тип, например, из-за опечатки в формуле */
            for (int i = 0; i < variablesTemp.Length; i++)
            {
                if(typeOfVariables[i] == TypeOfVariable.CELL)
                {
                    //TODO: Read from Excel
                    //BUT for test this values are 1
                    values[i] = 1;
                }
                else if (typeOfVariables[i] == TypeOfVariable.COEFFICIENT)
                {
                    //TODO: Insert calculated coefficient
                    //BUT for test this values are 2
                    values[i] = 2;
                }
                else if (typeOfVariables[i] == TypeOfVariable.CONST)
                {
                    double tempVarForConstantTryParse = 0;
                    Double.TryParse(variablesTemp[i], out tempVarForConstantTryParse);
                    values[i] = tempVarForConstantTryParse;
                }
                else if (typeOfVariables[i] == TypeOfVariable.UNKNOW)
                {                                       
                    values[i] = (int)ErrorCode.UNKNOW_VARIABLE_TYPE;
                    savesIntoFile.SaveIntoFile($"Ошибка в формуле {formula}: {variablesTemp[i]} имеет тип {typeOfVariables[i]}. Значению в формуле было присвоено {(int)ErrorCode.UNKNOW_VARIABLE_TYPE}", DateTime.Now.ToString("yyyyMMdd.HHmmss"), extensionOfErrorLogFile, false);
                }
            }

            //Строка, содержащая формулу, содержащая значения вместо имен ячеек. Должна высчитыватся в парсере математических выражений
            string formulaWithValues = GetFormulaWithValues(new StringBuilder(formula, formula.Length * 2), variablesTemp, values);

            double dResult = mathParser.Parse(formulaWithValues, isRadians);
            string sResult = Math.Round(dResult, 2).ToString();

            return sResult;        
        }

        private static string TEST_MakeCalculation(string formula, string source)
        {
            //Получение имени текущего метода для вывода в логи
            string currMethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

            MathParser mathParser = new MathParser();
            bool isRadians = false;

            //Сепараторы для парсинга формулы
            string[] stringSeparators = new string[] { "+", "-", "/", "*", "(", ")" };

            //Временный массив для хранения распарсенной формулы
            string[] variablesTemp = formula.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

            //Получение типов переменных формулы
            TypeOfVariable[] typeOfVariables = GetTypesOfVariables(variablesTemp);

            //Массив для значений переменных
            double[] values = new double[variablesTemp.Length];

            /*В соответствии с типом переменной для неё выполняются соответсвующе действия:
             * для TypeOfVariable.CELL значение считывается из соответствующей ячейки
             * для TypeOfVariable.COEFFICIENT подставляется расчитанное знчение коэффициента
             * для TypeOfVariable.CONST значение берется из массива переменных, которое передеется как аргумент в этот метод
             * тип TypeOfVariable.UNKNOW присваивается всем переменным, которые не попали ни под один другой тип, например, из-за опечатки в формуле */
            for (int i = 0; i < variablesTemp.Length; i++)
            {
                if (typeOfVariables[i] == TypeOfVariable.CELL)
                {
                    //TODO: Read from Excel
                    //BUT for test this values are 1
                    values[i] = 1;
                }
                else if (typeOfVariables[i] == TypeOfVariable.COEFFICIENT)
                {
                    //TODO: Insert calculated coefficient
                    //BUT for test this values are 2
                    values[i] = 2;
                }
                else if (typeOfVariables[i] == TypeOfVariable.CONST)
                {
                    double tempVarForConstantTryParse = 0;
                    if (Double.TryParse(variablesTemp[i], out tempVarForConstantTryParse))
                        values[i] = tempVarForConstantTryParse;
                    else
                    {
                        values[i] = ConvertDecimalSeparator(variablesTemp[i]);
                    }
                }
                else if (typeOfVariables[i] == TypeOfVariable.UNKNOW)
                {
                    values[i] = (int)ErrorCode.UNKNOW_VARIABLE_TYPE;
                    //savesIntoFile.SaveIntoFile($"Ошибка в формуле \"{formula}\": \"{variablesTemp[i]}\" имеет тип {typeOfVariables[i]}. Значению в формуле было присвоено {(int)ErrorCode.UNKNOW_VARIABLE_TYPE}", DateTime.Now.ToString("yyyyMMdd.HHmmss"), extensionOfErroeLogFile, false);
                    savesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{currMethodName}\r\nОшибка в формуле \"{formula}\": \"{variablesTemp[i]}\" имеет тип {typeOfVariables[i]}. Значению в формуле было присвоено {(int)ErrorCode.UNKNOW_VARIABLE_TYPE}\r\n", nameOfErrorLogFile, extensionOfErrorLogFile, true);
                }
            }

            //Строка, содержащая формулу, содержащая значения вместо имен ячеек. Должна высчитыватся в парсере математических выражений
            string formulaWithValues = GetFormulaWithValues(new StringBuilder(formula, formula.Length * 2), variablesTemp, values);

            double dResult = mathParser.Parse(formulaWithValues, isRadians);
            string sResult = Math.Round(dResult, 2).ToString();

            return sResult;
        }

        private static SeparatedCellNameStruct SeparateCellName(string cell)
        {
            SeparatedCellNameStruct separatedCellName = new SeparatedCellNameStruct();

            if (ExcelUtility.CheckData.IsCell(cell))
            {
                Regex regex = new Regex(@"(^[A-Z]{1,})([0-9]{1,}$)");

                Match match = regex.Match(cell);

                if(match.Success)
                {
                    separatedCellName.Letter = match.Groups[1].Value;
                    separatedCellName.Number = Convert.ToInt16(match.Groups[2].Value);
                }

            }
            else
                throw new System.FormatException();

            return separatedCellName;
        }

        static void Main(string[] args)
        {
            //string result = MakeCalculation("A1+A2");
            string result = MakeCalculation("1,047");
            Console.WriteLine(result);

            //string startRange = "A3";
            //string endRange = "D6";

            //SeparatedCellNameStruct separatedCellNameOfStartRange = SeparateCellName(startRange);
            //SeparatedCellNameStruct separatedCellNameOfEndRange = SeparateCellName(endRange);

            string a = "125,01";
            string b = "125.01";

            Console.WriteLine($"{a} is number: {ExcelUtility.CheckData.IsNumber(a)}");
            Console.WriteLine($"{b} is number: {ExcelUtility.CheckData.IsNumber(b)}");

            #region FormatException
            double dTest1 = 0;
            string sTest1 = "12.u5";

            try
            {
                dTest1 = Double.Parse(sTest1); //SystemFormat.Exception
                Console.WriteLine($"dTest1={dTest1} //try");
            }
            catch (System.FormatException)
            {
                Console.WriteLine("FormatException");

                try
                {
                    dTest1 = ConvertDecimalSeparator(sTest1);
                    Console.WriteLine($"dTest1={dTest1} //catch");
                }
                catch (System.FormatException)
                {
                    Console.WriteLine($"Parse Double Failure. {sTest1} is NaN");
                }
            }
            #endregion

            string sTest2 = "coeffK1";
            Console.WriteLine($"{sTest2} is coefficient: {ExcelUtility.CheckData.IsCoefficients(sTest2)}");

            string[] sTest3 = { "A1", "AA1", "AA11", "AB1", "AB11", "11", "11AB" };

            foreach(string elem in sTest3)
            {
                Console.WriteLine($"{elem} is cell: {ExcelUtility.CheckData.IsCell(elem)}");
            }

            string formula = "1,1/1.1/M20/coeffK1/1q1/M20000/coefK1/coeff K1/M20";
            Console.WriteLine($"{formula} is coefficient: {ExcelUtility.CheckData.IsCoefficients(formula)}");
            Console.WriteLine($"{formula} is cell: {ExcelUtility.CheckData.IsCell(formula)}");

            ////Сепараторы для парсинга формулы
            string[] stringSeparators = new string[] { "+", "-", "/", "*", "(", ")" };

            ////Временный массив для хранения распарсенной формулы
            string[] variablesTemp = formula.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

            TypeOfVariable[] typeOfVar = GetTypesOfVariables(variablesTemp);

            Console.WriteLine($"\r\n{formula}\r\n-----------\r\nElement\tType");
            for (int i = 0; i < variablesTemp.Length; i++)
            {
                Console.WriteLine($"{variablesTemp[i]}\t{typeOfVar[i]}");
            }

            //savesIntoFile.SaveIntoFile(TypeOfVariable.UNKNOW.ToString(), DateTime.Now.ToString("yyyyMMdd.HHmmss"), "err", false);
            string currMethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            //Console.WriteLine(currMethodName + "\t" + this.GetType().Name);
            Console.WriteLine((int)ErrorCode.UNKNOW_VARIABLE_TYPE);

            Console.WriteLine($"TEST_MakeCalculation(A1 + A2 + 1,25 + 45.7447 + coeffK1) = { TEST_MakeCalculation("A1+A2+1,25+45.7а447+coeffK1", "")}");

            string test4 = "A1+A2+1,25+45,7447+coeffK1";
            string[] stringSeparators1 = new string[] { "+", "-", "/", "*", "(", ")" };
            string[] variablesTemp1 = test4.Split(stringSeparators1, StringSplitOptions.RemoveEmptyEntries);
            //string[] variablesTemp = test4.Split(stringSeparators, StringSplitOptions.None);
            Console.WriteLine($"***************************************************");
            foreach (string elem in variablesTemp1)
                Console.WriteLine($"{elem} {Array.IndexOf(variablesTemp1, elem)}");

            string test5 = " sefgerher ";
            Console.WriteLine($"_{test5}_");
            test5 = test5.Trim();
            Console.WriteLine($"_{test5}_");

            string test6 = "BD132:К:K115/1000*K1*K3:Смета ТС/const*coeffK1*coeffK3";
            Marking marking = new Marking(test6);

            string cellForTest = "A1A";
            Console.WriteLine($"{cellForTest} is cell: {ExcelUtility.CheckData.IsCell(cellForTest)}");
            Console.WriteLine($"{SeparateCellName(cellForTest).Letter} {SeparateCellName(cellForTest).Number}");
        }
    }
}
