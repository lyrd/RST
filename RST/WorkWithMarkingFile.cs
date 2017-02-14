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
        //TODO: пути к файлам в ini
        private static readonly string heatSupplyMarkings = Directory.GetCurrentDirectory() + @"\Markings\heatSupplyMarking.txt";
        private static readonly string waterSupplyMarkings = Directory.GetCurrentDirectory() + @"\Markings\waterSupplyMarking.txt";
        private static readonly string sewerageMarkings = Directory.GetCurrentDirectory() + @"\Markings\sewerageMarking.txt";

        private static readonly string heatSupplyCoefficients = Directory.GetCurrentDirectory() + @"\Coefficients\heatSupplyCoefficients.txt";
        private static readonly string waterSupplyCoefficients = Directory.GetCurrentDirectory() + @"\Coefficients\waterSupplyCoefficients.txt";
        private static readonly string sewerageCoefficients = Directory.GetCurrentDirectory() + @"\Coefficients\sewerageCoefficients.txt";

        private string pathToMarkingFile = String.Empty;
        private string pathToCoefficientFile = String.Empty;

        public WorkWithMarkingFile()
        {

        }

        public void GetSourcesOfVariables()
        {

        }
    }
}
