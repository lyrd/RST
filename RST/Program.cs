using MathExpressionsParser;
using System;
using System.Collections.Generic;
using System.IO;
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
        private static MainLogic mainLogic = new MainLogic();

        private enum CVErrEnum : int
        {
            ErrDiv0 = -2146826281,
            ErrNA = -2146826246,
            ErrName = -2146826259,
            ErrNull = -2146826288,
            ErrNum = -2146826252,
            ErrRef = -2146826265,
            ErrValue = -2146826273
        }

        private static bool IsXLCVErr(object obj)
        {
            return (obj) is Int32;
        }

        private static bool IsXLCVErr(object obj, CVErrEnum whichError)
        {
            return (obj is Int32) && ((Int32)obj == (Int32)whichError);
        }

        static void Main(string[] args)
        {
            #region test
            //string result = MakeCalculation("A1+A2");
            //string result = mainLogic.MakeCalculation("1,047");
            //Console.WriteLine(result);

            //string startRange = "A3";
            //string endRange = "D6";

            //SeparatedCellNameStruct separatedCellNameOfStartRange = SeparateCellName(startRange);
            //SeparatedCellNameStruct separatedCellNameOfEndRange = SeparateCellName(endRange);

            //string a = "125,01";
            //string b = "125.01";

            //Console.WriteLine($"{a} is number: {ExcelUtility.CheckData.IsNumber(a)}");
            //Console.WriteLine($"{b} is number: {ExcelUtility.CheckData.IsNumber(b)}");

            #region FormatException
            //double dTest1 = 0;
            //string sTest1 = "12.u5";

            //try
            //{
            //    dTest1 = Double.Parse(sTest1); //SystemFormat.Exception
            //    //Console.WriteLine($"dTest1={dTest1} //try");
            //}
            //catch (System.FormatException)
            //{
            //    //Console.WriteLine("FormatException");

            //    try
            //    {
            //        dTest1 = mainLogic.ConvertDecimalSeparator(sTest1);
            //        //Console.WriteLine($"dTest1={dTest1} //catch");
            //    }
            //    catch (System.FormatException)
            //    {
            //        //Console.WriteLine($"Parse Double Failure. {sTest1} is NaN");
            //    }
            //}
            #endregion

            //string sTest2 = "coeffK1";
            // Console.WriteLine($"{sTest2} is coefficient: {ExcelUtility.CheckData.IsCoefficients(sTest2)}");

            //string[] sTest3 = { "A1", "AA1", "AA11", "AB1", "AB11", "11", "11AB" };

            //foreach(string elem in sTest3)
            //{
            //    //Console.WriteLine($"{elem} is cell: {ExcelUtility.CheckData.IsCell(elem)}");
            //}

            //string formula = "1,1/1.1/M20/coeffK1/1q1/M20000/coefK1/coeff K1/M20";
            //Console.WriteLine($"{formula} is coefficient: {ExcelUtility.CheckData.IsCoefficients(formula)}");
            //Console.WriteLine($"{formula} is cell: {ExcelUtility.CheckData.IsCell(formula)}");

            ////Сепараторы для парсинга формулы
            //string[] stringSeparators = new string[] { "+", "-", "/", "*", "(", ")" };

            ////Временный массив для хранения распарсенной формулы
            //string[] variablesTemp = formula.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

            //MainLogic.TypeOfVariable[] typeOfVar = MainLogic.GetTypesOfVariables(variablesTemp);

            //Console.WriteLine($"\r\n{formula}\r\n-----------\r\nElement\tType");
            //for (int i = 0; i < variablesTemp.Length; i++)
            //{
            //    //Console.WriteLine($"{variablesTemp[i]}\t{typeOfVar[i]}");
            //}

            //savesIntoFile.SaveIntoFile(TypeOfVariable.UNKNOW.ToString(), DateTime.Now.ToString("yyyyMMdd.HHmmss"), "err", false);
            //string currMethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            //Console.WriteLine(currMethodName + "\t" + this.GetType().Name);
            //Console.WriteLine((int)MainLogic.ErrorCode.UNKNOW_VARIABLE_TYPE);

            //Console.WriteLine($"TEST_MakeCalculation(A1 + A2 + 1,25 + 45.7447 + coeffK1) = { mainLogic.TEST_MakeCalculation("A1+A2+1,25+45.74.47+coeffK1", "")}");

            //string test4 = "A1+A2+1,25+45,7447+coeffK1";
            //string[] stringSeparators1 = new string[] { "+", "-", "/", "*", "(", ")" };
            //string[] variablesTemp1 = test4.Split(stringSeparators1, StringSplitOptions.RemoveEmptyEntries);
            //string[] variablesTemp = test4.Split(stringSeparators, StringSplitOptions.None);
            //Console.WriteLine($"***************************************************");
            //foreach (string elem in variablesTemp1)
            //    Console.WriteLine($"{elem} {Array.IndexOf(variablesTemp1, elem)}");

            //string test5 = " sefgerher ";
            //Console.WriteLine($"_{test5}_");
            //test5 = test5.Trim();
            //Console.WriteLine($"_{test5}_");

            //Console.WriteLine("*****************MARKING_TEST***********************");
            //string test6 = "BD132:К:K115/1000*K1*K3:Смета ТС/const*coeffK1*coeffK3";
            //GetMarkingTurple marking = new GetMarkingTurple(test6);
            //Console.WriteLine($"{test6}\r\nTarget Cell: {marking.TargetCell}\r\nTaget Sheet: {marking.TargetSheet}\r\nSource Cell: {marking.SourceCell}\r\nSource Sheet: {marking.SourceSheet}");

            //string cellForTest = "AB12";
            //Console.WriteLine($"{cellForTest} is cell: {ExcelUtility.CheckData.IsCell(cellForTest)}");
            //try
            //{
            //    Console.WriteLine($"Letter: {ExcelUtility.GetSeparatedCellName(cellForTest).Letter}\r\nNumber: {ExcelUtility.GetSeparatedCellName(cellForTest).Number}");
            //}
            //catch ( Exceptions.CellNameException cellEx )
            //{
            //    Console.WriteLine(cellEx.Message + "\r\n" + cellEx.StackTrace);
            //}

            //WorkWithMarkingFile wwmf = new WorkWithMarkingFile();
            //List<Marking> listOfMarking = wwmf.GetListOfMarkings();
            #endregion

            Office.Excel excel = new Office.Excel();
            excel.OpenDocument(Directory.GetCurrentDirectory() + @"\test\test_read.xlsx");

            for (int i = 1; i < 13; i++)
            {
                string sRead = excel.GetValue($"A{i}", "read");
                double dRead = 0;
                int iRead = 0;



                Console.WriteLine($"A{i}=\"{iRead}\"\t{IsXLCVErr(iRead, CVErrEnum.ErrDiv0)}");
            }

            excel.CloseDocument();
            excel.Dispose();

            Console.WriteLine(IsXLCVErr(-2146826281, CVErrEnum.ErrDiv0));
            Console.WriteLine(IsXLCVErr(-2146826281));

            Console.WriteLine(Enum.IsDefined(typeof(CVErrEnum), -2046826281));

            //OfficeLib.ExcelData excel = new OfficeLib.ExcelData(Directory.GetCurrentDirectory() + @"\test\test_read.xlsx");
            //for (int i = 1; i < 11; i++)
            //{
            //    string read = excel.Read($"A{i}", "read");
            //    Console.WriteLine($"A{i}=\"{read}\"\t{IsXLCVErr(read, CVErrEnum.ErrDiv0)}");
            //}
            //Console.WriteLine(excel.State);
            //excel.Close();
            //Console.WriteLine(excel.State);

            //Console.ReadKey();
        }
    }
}
