
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
//using System.Text.Json;

namespace RBExcelTool.Lib
{
    class ExcelData
    {
        public string mExcelPath;
        public string mSheetName;
        public string mExcelName;
    }
    internal class RBExcelHandler
    {
        private ExcelData mSameSheetName1, mSameSheetName2;
        private string mSheetName;
        private string mExcelPath;
        private Dictionary<string, ExcelData> mExcelDictionary;
        public RBExcelHandler(string _ExcelRootPath, string findSheetName)
        {
            mSheetName = findSheetName;
            //获取应用程序的当前工作目录
            string path3 = System.IO.Directory.GetCurrentDirectory();
            //var path = path3 + @"\Excel";
            //var path = path3 + @"\SVN";
            //mExcelPath = path;
            mExcelPath = _ExcelRootPath;
            _ = WriteJsonPathAsync();
            mExcelDictionary = new Dictionary<string, ExcelData>();
        }
        public void Process()
        {
            ConsoleColor ConsoleNormalColor = Console.ForegroundColor;
            Stopwatch st = new Stopwatch();
            st.Start();
            if (!Directory.Exists(mExcelPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("该路径【{0}】不存在，请重新指定 Excle 路径", mExcelPath);
                return;
            }
            mExcelDictionary.Clear();
            ResdJson();
            ExcelData _ExcelData;
            if (mExcelDictionary.TryGetValue(mSheetName, out _ExcelData))
            {
                st.Stop();
                TimeSpan ts = st.Elapsed;
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                Console.ForegroundColor = ConsoleNormalColor;
                Console.WriteLine("共用时【{0}】", elapsedTime);
                Console.WriteLine("【{0}】 所在的 Excel =【{1}】\n按回车键打开,按 ESC 键退出", mSheetName, _ExcelData.mExcelName);
                while (true)
                {
                    var cki = Console.ReadKey(true);
                    if (cki.Key == ConsoleKey.Enter)
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(_ExcelData.mExcelPath);
                        }
                        catch (System.ComponentModel.Win32Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(ex.Message);
                            Console.WriteLine("被打开的 Excel = 【{0}】", _ExcelData.mExcelPath);
                            Console.WriteLine("该 Excel 或被删除 或已被移动到其他文件夹内 按 回车键 重新查找 按 ESC 键退出");
                            while (true)
                            {
                                var _cki = Console.ReadKey(true);
                                if (_cki.Key == ConsoleKey.Enter)
                                {
                                    st = new Stopwatch();
                                    st.Start();
                                    goto A;
                                }
                                else if (cki.Key == ConsoleKey.Escape)
                                {
                                    break;
                                }
                                else
                                {
                                    Thread.Sleep(1);
                                }
                            }
                        }
                        break;
                    }
                    else if (cki.Key == ConsoleKey.Escape)
                    {
                        break;
                    }
                    else
                    {
                        Thread.Sleep(1);
                    }
                }
                return;
            }

        A: int ExcelCount = 0;
            Console.ForegroundColor = ConsoleNormalColor;
            mExcelDictionary.Clear();
            string[] files = Directory.GetFiles(mExcelPath, "*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var path = files[i];
                if (path.Contains("~$"))
                {
                    continue;
                }
                ExcelCount = i;
                Console.WriteLine(path);

                ProcessExcel(path);
            }

            if (mExcelDictionary.TryGetValue(mSheetName, out _ExcelData))
            {
                st.Stop();
                TimeSpan ts = st.Elapsed;
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                Console.WriteLine("共用时【{0}】", elapsedTime);
                Console.WriteLine("【{0}】 所在的 Excel =【{1}】\n按回车键打开,按下 ESC 键退出", mSheetName, _ExcelData.mExcelName);
                while (true)
                {
                    var cki = Console.ReadKey(true);
                    if (cki.Key == ConsoleKey.Enter)
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(_ExcelData.mExcelPath);
                        }
                        catch (System.ComponentModel.Win32Exception e)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Caught: {0}", e.Message);
                        }
                        catch (ObjectDisposedException e)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Caught: {0}", e.Message);
                        }
                        catch(PlatformNotSupportedException e)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Caught: {0}", e.Message);
                        }
                        break;
                    }
                    else if (cki.Key == ConsoleKey.Escape)
                    {
                        break;
                    }
                    else
                    {
                        Thread.Sleep(1);
                    }
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("未找到 【{0}】", mSheetName);
                st.Stop();
                TimeSpan ts = st.Elapsed;
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                Console.ForegroundColor = ConsoleNormalColor;
                Console.WriteLine("共用时【{0}】", elapsedTime);
            }
            Console.ForegroundColor = ConsoleNormalColor;
            Console.WriteLine("共有【{0}】个 Excel 文件", ExcelCount);
            _ = WriteJsonAsync();
        }
        bool ProcessExcel(string path)
        {
            var fileInfo = new FileInfo(path);
            var _ExcelName = Path.GetFileNameWithoutExtension(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                var count = excelPackage.Workbook.Worksheets.Count;
                for (int i = 0; i < count; i++)
                {
                    var worksheet = excelPackage.Workbook.Worksheets[i];
                    var sheetName = worksheet.Name;
                    {
                        ExcelData ExcelData;
                        if (!mExcelDictionary.TryGetValue(sheetName, out ExcelData))
                        {
                            ExcelData _ExcelData = new ExcelData
                            {
                                mExcelPath = path,
                                mExcelName = _ExcelName,
                                mSheetName = sheetName,
                            };
                            mExcelDictionary.Add(sheetName, _ExcelData);
                        }
                        else
                        {
                            mSameSheetName1 = new ExcelData
                            {
                                mExcelPath = path,
                                mExcelName = _ExcelName,
                                mSheetName = sheetName,
                            };
                            mSameSheetName2 = ExcelData;
                            //这里应当 结束 所有正在执行的异步方法
                            //return;
                        }
                    }
                }
            }
            return false;
        }

        async Task WriteJsonAsync()
        {
            byte[] info = new UTF8Encoding(true).GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(mExcelDictionary));
            string jsonPath = mExcelPath + "Excel.json";
            using (FileStream SourceStream = File.Open(jsonPath, FileMode.Create))
            {
                SourceStream.Seek(0, SeekOrigin.End);
                await SourceStream.WriteAsync(info, 0, info.Length);
            }
        }
        void ResdJson()
        {
            string jsonPath = mExcelPath + "Excel.json";
            if (File.Exists(jsonPath))
            {
                using (StreamReader file = File.OpenText(jsonPath))
                {
                    var ss = file.ReadToEnd();
                    var ExcelObj = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, ExcelData>>(ss);
                    if (ExcelObj != null)
                    {
                        mExcelDictionary = ExcelObj;
                    }
                }
            }
        }
        async Task WriteJsonPathAsync()
        {
            byte[] info = new UTF8Encoding(true).GetBytes(mExcelPath);
            string jsonPath = System.IO.Directory.GetCurrentDirectory();
            jsonPath = jsonPath + @"\ExcelRootPath.txt";
            using (FileStream SourceStream = File.Open(jsonPath, FileMode.Create))
            {
                SourceStream.Seek(0, SeekOrigin.End);
                await SourceStream.WriteAsync(info, 0, info.Length);
            }
        }
    }
}
