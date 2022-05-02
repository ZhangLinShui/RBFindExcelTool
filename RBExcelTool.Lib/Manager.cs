using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RBExcelTool.Lib
{
    public static class Manager
    {
        public static void Exprot(string _ExcelRootPath,string _SheetName)
        {
            new RBExcelHandler(_ExcelRootPath,_SheetName).Process();

        }
    }
}
