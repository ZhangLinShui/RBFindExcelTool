using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RBExcelTool.Lib
{
    public static class Manager
    {
       public static Regex searchPattern = new Regex(@"$(?<=\.(xlsx|xlsm|xlx))", RegexOptions.IgnoreCase);
        public static void Exprot(string _ExcelRootPath,string _SheetName)
        {
            new RBExcelHandler(_ExcelRootPath,_SheetName).Process();

        }
    }
}
