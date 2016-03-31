using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtentions.Samples {
    class Program {
        static void Main(string[] args) {
            new ExcelHelpTests().GetByDt();
            new ExcelHelpTests().GetByDs();
            new ExcelHelpTests().GetDynamic();
            new ExcelHelpTests().GetByDynamicList();
        }
    }
}
