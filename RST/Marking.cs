using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RST
{
    class Marking
    {
        //BD132:К:K115/1000*K1*K3:Смета ТС/const*coeffK1*coeffK3
        private string targetCell = String.Empty;  //BD132
        private string targetSheet = String.Empty; //К
        private string sourceCell = String.Empty;  //K115/1000*K1*K3
        private string sourceSheet = String.Empty; //Смета ТС/const*coeffK1*coeffK3        

        public string TargetCell
        {
            get { return this.targetCell; }
            //set { this.targetCell = value; }
        }

        public string TargetSheet
        {
            get { return this.targetSheet; }
            //set { this.targetSheet = value; }
        }

        public string SourceCell
        {
            get { return this.sourceCell; }
            set { this.sourceCell = value; }
        }

        public string SourceSheet
        {
            get { return this.sourceSheet; }
            set { this.sourceSheet = value; }
        }

        public Marking(string targetCell, string targetSheet, string sourceCell, string sourceSheet)
        {
            this.targetCell = targetCell;
            this.targetSheet = targetSheet;
            this.sourceCell = sourceCell;
            this.sourceSheet = sourceSheet;
        }

    }
}
