using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RST
{
    static class ExcelUtility
    {
        public static class CheckData
        {
            public static bool IsCell(string value)
            {
                string pattern = @"^[A-Z]{1,}[0-9]{1,}$";

                return Regex.IsMatch(value, pattern);
            }

            public static bool IsNumber(string value)
            {
                string pattern = @"^[0-9]{1,}[.,]{0,1}[0-9]{0,}$";

                return Regex.IsMatch(value, pattern);
            }

            public static bool IsCoefficients(string value)
            {
                string pattern = @"^coeffK[0-9]$";

                return Regex.IsMatch(value, pattern);
            }
        }

        public struct SeparatedCellName
        {
            private string letter;
            private int number;

            public string Letter
            {
                get { return this.letter; }
                set { this.letter = value; }
            }

            public int Number
            {
                get { return this.number; }
                set { this.number = value; }
            }
        }

        public static int GetColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        public static string GetColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int mod;

            while (dividend > 0)
            {
                mod = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + mod).ToString() + columnName;
                dividend = (int)((dividend - mod) / 26);
            }

            return columnName;
        }

        public static SeparatedCellName GetSeparatedCellName(string cell)
        {
            SeparatedCellName separatedCellName = new SeparatedCellName();

            if (ExcelUtility.CheckData.IsCell(cell))
            {
                Regex regex = new Regex(@"(^[A-Z]{1,})([0-9]{1,}$)");

                Match match = regex.Match(cell);

                if (match.Success)
                {
                    separatedCellName.Letter = match.Groups[1].Value;
                    separatedCellName.Number = Convert.ToInt16(match.Groups[2].Value);
                }

            }
            else
                throw new Exceptions.CellNameException($"Аргумент метода {System.Reflection.MethodBase.GetCurrentMethod().Name} не является ячейкой");

            return separatedCellName;
        }
    }
}
