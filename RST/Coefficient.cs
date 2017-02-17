using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RST
{
    class Coefficient
    {
        private double coeff = 0;
        private string formula = String.Empty;

        private static readonly string heatSupplyCoefficients = Directory.GetCurrentDirectory() + @"\Coefficients\heatSupplyCoefficients.txt";
        private static readonly string waterSupplyCoefficients = Directory.GetCurrentDirectory() + @"\Coefficients\waterSupplyCoefficients.txt";
        private static readonly string sewerageCoefficients = Directory.GetCurrentDirectory() + @"\Coefficients\sewerageCoefficients.txt";

        public double Coeff
        {
            get { return this.coeff; }
        }


        public Coefficient(string formula)
        {
            this.formula = formula;
        }

        public void MakeCalculation()
        {

        }
    }
}
