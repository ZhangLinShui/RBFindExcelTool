using RBExcelTool.Lib;
using System;
using System.IO;
using System.Text;

namespace RBFindExcelTool
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string _ExcelRootPath = null;
            string jsonPath = System.IO.Directory.GetCurrentDirectory();
            jsonPath = jsonPath + @"\ExcelRootPath.txt";
            string tips = "首次运行需要指定 Excel 所在根目录";
            if (File.Exists(jsonPath))
            {
                using (StreamReader file = File.OpenText(jsonPath))
                {
                    _ExcelRootPath = file.ReadToEnd();
                }
                if (_ExcelRootPath != null || _ExcelRootPath != "")
                {
                    if (ChackPath(_ExcelRootPath) == true)
                    {
                        goto B;
                    }
                    else
                    {
                        tips = "Excel 所在目录更换后 需重新指定 根目录";
                        goto C;
                    }
                }
            }
        C: Console.WriteLine(tips);
        A: Console.WriteLine("请输入需要查找的 Excel 所在根目录");

            _ExcelRootPath = Console.ReadLine();
            if (_ExcelRootPath.Contains("\n") || string.IsNullOrEmpty(_ExcelRootPath) || _ExcelRootPath.Contains(" ") || _ExcelRootPath.Contains("\r") ||
                _ExcelRootPath.Contains("\r\n"))
            {
                Console.WriteLine("输入路径非法 请重新输入");
                goto A;
            }
            if (ChackPath(_ExcelRootPath) == false)
            {
                goto A;
            }

        B: Console.WriteLine("需要查找的 Excel 所在根目录 = 【{0}】", _ExcelRootPath);

        D: Console.WriteLine("请输入需要查找的 SheetName");
            string _SheetName = Console.ReadLine();
            Console.WriteLine("需要查找的 SheetName = 【{0}】", _SheetName);
            if (_SheetName.Contains("\n") || string.IsNullOrEmpty(_SheetName)|| _SheetName.Contains(" ")|| _SheetName.Contains("\r")||
                _SheetName.Contains("\r\n"))
            {
                Console.WriteLine("输入 SheetName 非法 请重新输入");
                goto D;
            }

            Console.WriteLine("开始查找Excel");
            Manager.Exprot(_ExcelRootPath, _SheetName);
            Console.WriteLine("结束查找Excel");

            Console.ReadKey();
        }
        static bool ChackPath(string _ExcelRootPath)
        {
            if (Directory.Exists(_ExcelRootPath))
            {
                string[] files = Directory.GetFiles(_ExcelRootPath, "*.xlsx", SearchOption.AllDirectories);
                if (files.Length <= 0)
                {
                    Console.WriteLine("需要查找的 Excel 根目录 = 【{0}】\n该目录下，不存在任何 Excel 文件，请检查路径 并重新输入路径", _ExcelRootPath);
                    return false;
                }
            }
            else
            {
                Console.WriteLine("需要查找的 Excel 根目录 = 【{0}】\n该目录不存在，请检查路径 并重新输入路径", _ExcelRootPath);
                return false;
            }
            return true;
        }
    }
}
