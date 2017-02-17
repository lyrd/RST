using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MathExpressionsParser;

namespace RST
{
    class MainLogic
    {
        public MainLogic(){}

        private MathParser mathParser = new MathParser();

        private enum Delta
        {
            NDS = 1,
            K = 15,
            T = 15,
            BPr = 28,
            TM2_WATER = 10,
            TM2_WARM = 150
        }

        private enum Direction
        {
            RIGHT,
            DOWN
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

        //private
        public enum TypeOfVariable
        {
            CELL = 1,
            CONST = 2,
            COEFFICIENT = 3,
            UNKNOW = -1
        }

        //private
        public enum ErrorCode
        {
            UNKNOW_VARIABLE_TYPE = 1302201701,
            TRY_PARSE_FAILURE = 1302201702
        }

        //private non-static
        public static TypeOfVariable[] GetTypesOfVariables(string[] variablesInFormula)
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

        public string MakeCalculation(string formula)
        {
            //Получение имени текущего метода для вывода в логи
            string currMethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

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
                    SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{currMethodName}\r\nОшибка в формуле \"{formula}\": \"{variablesTemp[i]}\" имеет тип {typeOfVariables[i]}. Значению в формуле было присвоено {(int)ErrorCode.UNKNOW_VARIABLE_TYPE}\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                }
            }

            //Строка, содержащая формулу, содержащая значения вместо имен ячеек. Должна высчитыватся в парсере математических выражений
            string formulaWithValues = GetFormulaWithValues(new StringBuilder(formula, formula.Length * 2), variablesTemp, values);

            double dResult = mathParser.Parse(formulaWithValues, isRadians);
            string sResult = Math.Round(dResult, 2).ToString();

            return sResult;
        }

        public string TEST_MakeCalculation(string formula, string source)
        {
            //Получение имени текущего метода для вывода в логи
            string currMethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

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
                    SavesIntoFile.SaveIntoFile($"{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}\r\n{currMethodName}\r\nОшибка в формуле \"{formula}\": \"{variablesTemp[i]}\" имеет тип {typeOfVariables[i]}. Значению в формуле было присвоено {(int)ErrorCode.UNKNOW_VARIABLE_TYPE}\r\n", Constants.nameOfErrorLogFile, Constants.extensionOfErrorLogFile, true);
                }
            }

            //Строка, содержащая формулу, содержащая значения вместо имен ячеек. Должна высчитыватся в парсере математических выражений
            string formulaWithValues = GetFormulaWithValues(new StringBuilder(formula, formula.Length * 2), variablesTemp, values);

            double dResult = mathParser.Parse(formulaWithValues, isRadians);
            string sResult = Math.Round(dResult, 2).ToString();

            return sResult;
        }

        //private
        public string GetFormulaWithValues(StringBuilder data, string[] oldInString, double[] valuesForReplace)
        {
            for (byte index = 0; index < valuesForReplace.Length; index++)
            {
                data.Replace(oldInString[index], valuesForReplace[index].ToString());
            }

            return data.ToString();
        }

        //private
        public double ConvertDecimalSeparator(string s)
        {
            return Double.Parse(s.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
