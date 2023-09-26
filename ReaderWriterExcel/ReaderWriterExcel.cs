using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ReaderWriterExcel
{
    public class ReaderWriterExcel
    {
        public static Excel.Application xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        public Workbook book = xlApp.ActiveWorkbook;
        public static Worksheet input, result;
        public int heightBase, widthBase;

        public List<T> ReadFromExcel<T>(List<T> inDatas, int sheetNumber)
        {
            var worksheets = GetWorksheets();
            input = book.Sheets.Item[worksheets[sheetNumber]];
            heightBase = input.UsedRange.Rows.Count;
            widthBase = input.UsedRange.Columns.Count;

            var range = GetExcelData(input, 1, 1, heightBase, widthBase);
            var dataExcel = (object[,])range.Value2;
            if (dataExcel != null)
            {
                var fields = GetFields<T>()
                .Where(x => !x.Contains("<Id>"))
                .ToList();
                var fieldsCount = fields.Count();

                for (int i = 1; i < heightBase + 1; i++)
                {
                    var item = (T)Activator.CreateInstance(typeof(T));
                    for (int j = 1; j < widthBase + 1; j++)
                    {
                        if (dataExcel[i, j] != null)
                        {
                            string data = dataExcel[i, j].ToString()
                                .Trim();

                            item.GetType().GetField(fields.ElementAt(j - 1),BindingFlags.Instance |
                                BindingFlags.NonPublic)
                                .SetValue(item, data);
                        }
                    }

                    inDatas.Add(item);
                }
            }

            return inDatas;
        }

        public void SaveToExcel<T>(List<T> outDatas, int sheetNumber)
        {
            var worksheets = GetWorksheets();
            result = book.Sheets.Item[worksheets[sheetNumber]];

            var count = outDatas.Count();
            var fields = GetFields<T>()
            .Where(x => !x.Contains("<Id>"))
            .ToList();
            var fieldsCount = fields.Count();

            object[,] datas = new object[count, fieldsCount];

            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < fieldsCount; j++)
                {
                    datas[i, j] = outDatas[i].GetType().GetField(fields.ElementAt(j),
                        BindingFlags.Instance | BindingFlags.NonPublic)
                        .GetValue(outDatas[i])
                        .ToString();
                }
            }

            var range = GetExcelData(result, 1, 1, count, fieldsCount);
            range.Value = datas;
        }

        private static string[] GetFields<T>()
        {
            var fields = typeof(T).GetRuntimeFields()
                .Select(x => x.Name)
                .ToArray();

            return fields;
        }

        private string[] GetWorksheets()
        {
            var worksheets = new string[book.Sheets.Count];
            for (int i = 0; i < book.Sheets.Count; i++)
                worksheets[i] = book.Sheets[i + 1].Name;

            return worksheets;
        }

        private static Range GetExcelData(Worksheet input, int rightUp, int leftUp, int heightBase, int widthBase)
        {
            if (heightBase == 0)
                heightBase = rightUp;
            if (widthBase == 0)
                widthBase = leftUp;

            return input.Range[input.Cells[rightUp, leftUp], input.Cells[heightBase, widthBase]];
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        public static void ExcelKill()
        {
            GetWindowThreadProcessId(xlApp.Hwnd, out int activeExcelId);

            Process[] processList = Process
                .GetProcesses()
                .Where(name => name.ProcessName == "EXCEL")
                .Where(app => app.Id != activeExcelId)
                .ToArray();

            foreach (Process process in processList)
            {
                process.Kill();
                process.WaitForExit();
            }
        }
    }
}
