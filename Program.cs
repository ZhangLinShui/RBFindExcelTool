using RBExcelTool.Lib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

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
        A: Console.WriteLine("输入需要查找的 Excel 所在根目录");

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
            Console.WriteLine("如需更换 Excel 根目录 请按 回车键 继续请按 其他键");
            while (true)
            {
                var cki = Console.ReadKey(true);
                if (cki.Key == ConsoleKey.Enter)
                {
                    goto A;
                }
                else if (cki.Key != ConsoleKey.Enter)
                {
                    break;
                }
                else
                {
                    Thread.Sleep(1);
                }
            }

        D: Console.WriteLine("请输入需要查找的 SheetName");
            string _SheetName = Console.ReadLine();
            Console.WriteLine("需要查找的 SheetName = 【{0}】", _SheetName);
            if (_SheetName.Contains("\n") || string.IsNullOrEmpty(_SheetName) || _SheetName.Contains(" ") || _SheetName.Contains("\r"))
            {
                Console.WriteLine("输入 SheetName 非法 请重新输入");
                goto D;
            }
            Console.WriteLine("如需更换 SheetName 请按 回车键 继续请按 其他键");
            while (true)
            {
                var cki = Console.ReadKey(true);
                if (cki.Key == ConsoleKey.Enter)
                {
                    goto D;
                }
                else if (cki.Key != ConsoleKey.Enter)
                {
                    break;
                }
                else
                {
                    Thread.Sleep(1);
                }
            }

            Console.WriteLine("开始查找Excel");
            Manager.Exprot(_ExcelRootPath, _SheetName);
            Console.WriteLine("结束查找Excel");

            Console.WriteLine("按任意键清屏并继续");
            Console.ReadKey();
            Console.Clear();
            Console.WriteLine("如需继续查找 请按 回车键 退出请安 其他键");
            while (true)
            {
                var cki = Console.ReadKey(true);
                if (cki.Key == ConsoleKey.Enter)
                {
                    goto B;
                }
                else if (cki.Key != ConsoleKey.Enter)
                {
                    break;
                }
                else
                {
                    Thread.Sleep(1);
                }
            }
        }
        static bool ChackPath(string _ExcelRootPath)
        {
            if (Directory.Exists(_ExcelRootPath))
            {
                //string[] files = Directory.GetFiles(_ExcelRootPath, "*.xlsx", SearchOption.AllDirectories);
                //string[] files = Directory.GetFiles(_ExcelRootPath, "*.*", SearchOption.AllDirectories).Where(file => file.ToLower().EndsWith("xlsx") || file.ToLower().EndsWith("xlsm")).ToArray();
                //IEnumerable<string> extensions = new [] { "xlsx", "xlsm" };
                //var files = Directory.GetFiles(_ExcelRootPath, "*.*").Where(f => extensions.Contains(Path.GetExtension(f).ToLower()));


                
                var files = Directory.GetFiles(_ExcelRootPath).Where(f => Manager.searchPattern.IsMatch(f)).ToArray();



                if (files.Length<= 0)
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
