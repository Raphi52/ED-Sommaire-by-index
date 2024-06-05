using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EDSommaireByINDEX
{
  public class SummaryRow
    {
        public string sheet;
        public string Description1;
        public string Description2;
        public string function;
        public List<string> Revisions;
        public string prjNum;

        public SummaryRow()
        {
            Revisions = new List<string>();
        }
        public override string ToString()
        {
            return sheet + prjNum;
        }
    }
}
